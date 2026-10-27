// ==== G2G Migration Mapping Tool v2.2.0 — single-file install. Paste this whole file into Code.gs ====

// ================================================================ Main.gs
/**
 * Google-to-Google Migration Mapping Tool  (Main.gs: menu, settings, collection engine)
 * ===================================================================================
 * Runs inside the SOURCE tenant, signed in as its super admin. Everything happens in Google Sheets/Drive:
 *
 *   1. Collect source data  -> RAW_* tabs (Admin SDK, Groups Settings API, Drive API)
 *   2. Build review sheet   -> a separate client-facing spreadsheet (users, groups, members, shared drives, ACLs)
 *   3. Export files         -> one .zip in Drive: CloudM CSVs + files to rebuild the source structure at the destination
 *
 * The destination is a 1:1 rebuild of the source using placeholder addresses: user@domain -> user@temp.domain.
 * No GCP project, service account, domain-wide delegation or GAM. The script is owned by the tenant's own
 * admin, so it counts as an internal app and needs no OAuth verification.
 */

const TOOL_VERSION = '2.2.0';
const TIME_BUDGET_MS = 4.5 * 60 * 1000;   // Apps Script hard limit is 6 min per execution
const STATE_KEY = 'G2G_STATE';
const REVIEW_ID_KEY = 'G2G_REVIEW_ID';

const SETTINGS_TAB = 'Settings';
const SETTINGS = [
  // [key, label, default, help]
  ['project', 'Project / customer name', '', 'Used in file names and the review title'],
  ['tempPrefix', 'Temp subdomain prefix', 'temp.', 'user@domain.com -> user@temp.domain.com'],
  ['archivePrefix', 'Post-cutover source prefix', 'archive.', 'Source account address after cutover'],
  ['drivePrefix', 'Destination shared drive name prefix', '', 'Optional, e.g. "Boggs " -> "Boggs Admin"'],
  ['staleDays', 'Flag users with no sign-in for (days)', 90, 'Inactive users default to Migrate? = Review'],
  ['pwLength', 'Temporary password length', 16, 'Passwords are generated at export, never shown in the review'],
  ['domainCatchAll', 'Add domain catch-all to address replacements', 'Yes', 'Yes/No — see Read Me'],
  ['reviewUrl', 'Review spreadsheet URL', '', 'Filled by step 2. Replace with the approved copy if the client returned a different file'],
  ['status', 'Status', '', 'Updated automatically'],
  ['exportUrl', 'Last export (.zip) URL', '', 'Filled by step 3'],
];

const PHASES = ['runinfo', 'domains', 'orgunits', 'users', 'tenantusage', 'userusage', 'licenses', 'groups', 'members', 'settings',
                'drives', 'perms'];

// Admin SDK Reports API (usage). Data is dated in Pacific time and lags 1–3 days, so we walk back to the newest
// date that has data. Usage is optional: if the Reports API fails, collection continues without it.
const USAGE_LOOKBACK_DAYS = 7;
// Photos is deliberately not requested: we don't migrate it. "Quota used %" is Google's figure and still includes all Google storage.
const USER_USAGE_PARAMS = ['gmail_used_quota_in_mb', 'drive_used_quota_in_mb', 'total_quota_in_mb', 'used_quota_in_percentage'];
// used_quota_in_mb is requested only as the "data is ready for this date" signal; it includes Photos so it isn't displayed.
const TENANT_USAGE_PARAMS = ['used_quota_in_mb', 'total_quota_in_mb', 'gmail_used_quota_in_mb', 'drive_used_quota_in_mb',
  'team_drive_used_quota_in_mb', 'num_users', 'num_users_used_quota_ge_80percent', 'num_users_used_quota_ge_50_lt_80percent',
  'num_archived_users', 'num_suspended_users'];

// Enterprise License Manager. Workspace plans, Archived User, and Cloud Identity (users with no Workspace plan).
const LICENSE_PRODUCTS = ['Google-Apps', '101034', '101001', '101005'];

const RAW = {
  runinfo:  { name: 'RUN_Info', headers: ['key', 'value'] },
  domains:  { name: 'RAW_Domains', headers: ['domainName', 'type', 'isPrimary', 'verified', 'parentDomain'] },
  orgunits: { name: 'RAW_OrgUnits', headers: ['orgUnitPath', 'name', 'description', 'parentOrgUnitPath', 'orgUnitId'] },
  users:    { name: 'RAW_Users', headers: ['primaryEmail', 'givenName', 'familyName', 'fullName', 'orgUnitPath',
              'suspended', 'archived', 'isAdmin', 'isDelegatedAdmin', 'aliases', 'nonEditableAliases',
              'creationTime', 'lastLoginTime', 'isMailboxSetup', 'isEnrolledIn2Sv', 'customerId', 'id'] },
  groups:   { name: 'RAW_Groups', headers: ['email', 'name', 'description', 'directMembersCount', 'aliases',
              'nonEditableAliases', 'adminCreated', 'id'] },
  members:  { name: 'RAW_GroupMembers', headers: ['groupEmail', 'memberEmail', 'role', 'type', 'status',
              'delivery_settings', 'memberId'] },
  settings: { name: 'RAW_GroupSettings', headers: ['groupEmail', 'whoCanJoin', 'whoCanViewMembership',
              'whoCanViewGroup', 'whoCanPostMessage', 'allowExternalMembers', 'whoCanContactOwner',
              'whoCanDiscoverGroup', 'whoCanLeaveGroup', 'messageModerationLevel', 'spamModerationLevel',
              'whoCanModerateMembers', 'whoCanModerateContent', 'whoCanAssistContent', 'allowWebPosting',
              'isArchived', 'archiveOnly', 'membersCanPostAsTheGroup', 'includeInGlobalAddressList',
              'replyTo', 'customReplyTo', 'sendMessageDenyNotification', 'enableCollaborativeInbox',
              'primaryLanguage', 'error', 'rawJson'] },
  drives:   { name: 'RAW_SharedDrives', headers: ['driveId', 'name', 'createdTime', 'hidden', 'orgUnitId',
              'adminManagedRestrictions', 'copyRequiresWriterPermission', 'domainUsersOnly',
              'driveMembersOnly', 'sharingFoldersRequiresOrganizerPermission'] },
  tenantusage: { name: 'RAW_TenantUsage', headers: ['metric', 'value'] },
  userusage: { name: 'RAW_UserUsage', headers: ['email', 'gmailMB', 'driveMB', 'quotaMB', 'usedPct', 'reportDate'] },
  licenses: { name: 'RAW_Licenses', headers: ['email', 'productId', 'productName', 'skuId', 'skuName', 'error'] },
  perms:    { name: 'RAW_DrivePermissions', headers: ['driveId', 'driveName', 'type', 'role', 'emailAddress',
              'domain', 'displayName', 'deleted', 'permissionId', 'error'] },
};

// ------------------------------------------------------------------------------------------ menu

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Migration Mapping')
    .addItem('1. Collect source data', 'menuCollect')
    .addItem('2. Build review spreadsheet', 'menuBuild')
    .addItem('3. Export files for engineering', 'menuExport')
    .addSeparator()
    .addItem('Show status', 'showStatus')
    .addItem('Delete exported files (contain passwords)', 'menuDeleteExports')
    .addItem('Stop collection & reset', 'resetCollector')
    .addToUi();
  try { ensureSettings_(); } catch (e) { /* first open before authorization */ }
}

function menuCollect() {
  ensureSettings_();
  const ui = SpreadsheetApp.getUi();
  if (loadState_() && loadState_().status === 'RUNNING') {
    if (ui.alert('A collection is already running. Restart from scratch?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  startRun_();
}

function showStatus() {
  const s = loadState_();
  const msg = !s ? 'No collection has been run yet.'
    : `Collection: ${s.status}\nPhase: ${s.phaseIdx < PHASES.length ? PHASES[s.phaseIdx] : 'done'} ` +
      `(${Math.min(s.phaseIdx + 1, PHASES.length)}/${PHASES.length})\nStarted: ${s.startedAt}\nLast update: ${s.updatedAt}` +
      (s.lastError ? `\nLast error: ${s.lastError}` : '');
  SpreadsheetApp.getUi().alert(msg);
}

function resetCollector() {
  deleteContinueTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
  setSetting_('status', 'Collection reset');
  notify_('Collection stopped and reset.');
}

/** Time-based trigger target: resumes a checkpointed collection. */
function continueRun() {
  deleteContinueTriggers_();
  runLoop_();
}

// ------------------------------------------------------------------------------------------ settings

const INSTRUCTIONS = [
  ['How to use this tool (Project Managers)', 'title'],
  ['Before you start: sign in to the browser as the SOURCE tenant super admin (preferably a new profile in Chrome, but you can use an incognito/guest window). '],
  ['', ''],
  ['Step 0: Settings tab', 'h'],
  ['Fill in "Project / customer name". It is recommended to use the same name defined in HubSpot. Change the other settings only if the project uses a different convention ' +
   '(the defaults are user@temp.domain.com placeholders and user@archive.domain.com after cutover).', ''],
  ['', ''],
  ['Step 1: Migration Mapping menu → 1. Collect source data', 'h'],
  ['The first time, Google asks you to authorize the tool: choose the admin account → Allow. Then run step 1 again if it did not start. ' +
   'Small tenants take about a minute. Large ones continue in the background on their own; you can close the tab. ' +
   'Check the "Status" row on the Settings tab. It says "Collected …" when finished.', ''],
  ['', ''],
  ['Step 2: Migration Mapping menu → 2. Build review spreadsheet', 'h'],
  ['This creates a separate spreadsheet for the client that can be customized before sharing. Check the "Issues to Resolve" tab first and fix or explain anything you can. ' +
   'Then share it with the client approver (Share button), or File → Download → Microsoft Excel and email it.', ''],
  ['The client sets Migrate? to Yes or No on every row, corrects the yellow columns, and fills in the sign-off on the first tab.', ''],
  ['If the client sends back an Excel file instead of editing online: upload it to Drive, open it with Google Sheets, ' +
   'and paste its URL into Settings → "Review spreadsheet URL".', ''],
  ['', ''],
  ['Step 3: Migration Mapping menu → 3. Export files for engineering', 'h'],
  ['Download the .zip from the dialog and send it to the engineering team through a secure channel. It contains temporary passwords. ' +
   'Then run Migration Mapping → Delete exported files.', ''],
  ['', ''],
  ['If something changes in the tenant before cutover, run steps 1–3 again. The previous review is renamed [SUPERSEDED].', ''],
  ['When the project is finished, delete this copy of the tool from the customer tenant.', ''],
];

function ensureInstructions_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName('Instructions')) return;
  const sh = ss.insertSheet('Instructions', 0);
  sh.getRange(1, 1, INSTRUCTIONS.length, 1).setValues(INSTRUCTIONS.map(l => [l[0]])).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 900);
  INSTRUCTIONS.forEach((l, i) => {
    if (l[1] === 'title') sh.getRange(i + 1, 1).setFontSize(16).setFontWeight('bold');
    if (l[1] === 'h') sh.getRange(i + 1, 1).setFontWeight('bold').setFontSize(12).setBackground('#DDEBF7');
  });
}

function ensureSettings_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SETTINGS_TAB);
  if (sh) return sh;
  ensureInstructions_();
  sh = ss.insertSheet(SETTINGS_TAB, 1);
  sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]).setFontWeight('bold')
    .setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.getRange(2, 1, SETTINGS.length, 3).setValues(SETTINGS.map(s => [s[1], s[2], s[3]]));
  sh.getRange(2, 2, SETTINGS.length, 1).setNumberFormat('@').setBackground('#FFF2CC');
  const rowOf = k => SETTINGS.findIndex(s => s[0] === k) + 2;
  ['reviewUrl', 'status', 'exportUrl'].forEach(k => sh.getRange(rowOf(k), 2).setBackground('#EDEDED'));
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 380); sh.setColumnWidth(3, 520);
  sh.setFrozenRows(1);
  return sh;
}

function getSettings_() {
  const sh = ensureSettings_();
  const vals = sh.getRange(2, 1, SETTINGS.length, 2).getValues();
  const out = {};
  SETTINGS.forEach((s, i) => { out[s[0]] = String(vals[i][1] === '' ? s[2] : vals[i][1]).trim(); });
  // Sheets cells lose trailing spaces easily: "Boggs" -> "Boggs Admin", while "Boggs -" / "Boggs_" are kept as typed.
  if (out.drivePrefix && /[A-Za-z0-9]$/.test(out.drivePrefix)) out.drivePrefix += ' ';
  out.staleDays = parseInt(out.staleDays, 10) || 90;
  out.pwLength = Math.max(12, parseInt(out.pwLength, 10) || 16);
  if (!out.project) {
    const prim = readRaw_('domains').find(d => truthy_(d.isPrimary));
    out.project = prim ? prim.domainName : 'Migration';
  }
  return out;
}

function setSetting_(key, value) {
  const sh = ensureSettings_();
  sh.getRange(SETTINGS.findIndex(s => s[0] === key) + 2, 2).setValue(value);
}

// ------------------------------------------------------------------------------------------ collection engine

function startRun_() {
  deleteContinueTriggers_();
  const ss = SpreadsheetApp.getActive();
  Object.keys(RAW).forEach(k => { const sh = ss.getSheetByName(RAW[k].name); if (sh) ss.deleteSheet(sh); });
  saveState_({ phaseIdx: 0, cursor: {}, startedAt: new Date().toISOString(), status: 'RUNNING' });
  setSetting_('status', 'Collecting… (large tenants continue automatically every minute)');
  notify_('Collecting source data. You can close this tab — large tenants keep going in the background.');
  runLoop_();
}

function runLoop_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const state = loadState_();
    if (!state || state.status !== 'RUNNING') return;
    const deadline = Date.now() + TIME_BUDGET_MS;
    while (state.phaseIdx < PHASES.length) {
      const phase = PHASES[state.phaseIdx];
      if (!state.cursor.started) { initRawSheet_(phase); state.cursor.started = true; }
      let done;
      try {
        done = COLLECT[phase](state, deadline);
      } catch (e) {
        state.status = 'ERROR'; state.lastError = `${phase}: ${e.message}`;
        saveState_(state);
        setSetting_('status', `Collection FAILED in ${phase}: ${e.message}`);
        throw e;
      }
      if (done) { state.phaseIdx++; state.cursor = {}; }
      saveState_(state);
      if (!done) {
        setSetting_('status', `Collecting… phase ${state.phaseIdx + 1}/${PHASES.length} (${PHASES[state.phaseIdx]}), continuing automatically`);
        ScriptApp.newTrigger('continueRun').timeBased().after(60 * 1000).create();
        return;
      }
    }
    state.status = 'COMPLETE';
    appendRows_('runinfo', [['finishedAt', new Date().toISOString()], ['status', 'COMPLETE']]);
    saveState_(state);
    const ss = SpreadsheetApp.getActive();
    Object.keys(RAW).forEach(k => { const sh = ss.getSheetByName(RAW[k].name); if (sh) sh.hideSheet(); });
    const c = n => Math.max(0, (ss.getSheetByName(RAW[n].name).getLastRow() || 1) - 1);
    const usageInfo = {};
    readRaw_('tenantusage').forEach(r => { usageInfo[r.metric] = r.value; });
    const usageNote = usageInfo.status === 'OK' ? `storage as of ${usageInfo.reportDate}` : 'storage usage unavailable';
    setSetting_('status', `Collected ${new Date().toISOString().slice(0, 16)}Z: ${c('users')} users, ${c('groups')} groups, ` +
                          `${c('drives')} shared drives, ${usageNote}. Next: step 2.`);
    notify_('Collection complete. Next: Migration Mapping → 2. Build review spreadsheet.');
  } finally {
    lock.releaseLock();
  }
}

/** Reads a usage report's parameter list into {name: number|string}, dropping the "accounts:" prefix. */
function usageParams_(report) {
  const out = {};
  ((report && report.parameters) || []).forEach(p => {
    const k = String(p.name).replace(/^accounts:/, '');
    out[k] = p.intValue !== undefined ? Number(p.intValue) : (p.stringValue !== undefined ? p.stringValue : '');
  });
  return out;
}

function pacificDate_(daysAgo) {
  return Utilities.formatDate(new Date(Date.now() - daysAgo * 86400000), 'America/Los_Angeles', 'yyyy-MM-dd');
}

const COLLECT = {
  /** Finds the newest date with tenant usage data and stores tenant totals. Never fails the run. */
  tenantusage: function (state) {
    const params = TENANT_USAGE_PARAMS.map(p => 'accounts:' + p).join(',');
    const tried = [];
    for (let d = 1; d <= USAGE_LOOKBACK_DAYS; d++) {
      const date = pacificDate_(d);
      try {
        const res = call_(() => AdminReports.CustomerUsageReports.get(date, { parameters: params }));
        const vals = usageParams_((res.usageReports || [])[0]);
        if (vals.used_quota_in_mb === undefined || vals.used_quota_in_mb === '') { tried.push(`${date}: no data yet`); continue; }
        const rows = [['reportDate', date], ['status', 'OK']];
        TENANT_USAGE_PARAMS.forEach(p => rows.push([p, vals[p] === undefined ? '' : vals[p]]));
        appendRows_('tenantusage', rows);
        return true;
      } catch (e) {
        tried.push(`${date}: ${e.message}`);
        // Permission/disabled errors won't improve on older dates; stop early.
        if (/permission|not authorized|forbidden|access denied|disabled/i.test(e.message)) break;
      }
    }
    appendRows_('tenantusage', [['reportDate', ''], ['status', 'UNAVAILABLE'], ['detail', tried.join(' | ').slice(0, 45000)]]);
    return true;
  },

  /** Per-user storage for the date found by tenantusage. Never fails the run. */
  userusage: function (state, deadline) {
    const info = {};
    readRaw_('tenantusage').forEach(r => { info[r.metric] = r.value; });
    const date = info.reportDate;
    if (!date) return true;
    const params = USER_USAGE_PARAMS.map(p => 'accounts:' + p).join(',');
    try {
      return paginate_(state, deadline, token =>
        call_(() => AdminReports.UserUsageReport.get('all', date, { parameters: params, maxResults: 500, pageToken: token })),
        res => (res.usageReports || []).filter(r => r.entity && r.entity.userEmail).map(r => {
          const v = usageParams_(r);
          return [String(r.entity.userEmail).toLowerCase()].concat(USER_USAGE_PARAMS.map(p => v[p] === undefined ? '' : v[p]), [date]);
        }), 'userusage');
    } catch (e) {
      appendRows_('tenantusage', [['userUsageError', e.message]]);
      return true;
    }
  },

  runinfo: function (state) {
    appendRows_('runinfo', [['toolVersion', TOOL_VERSION], ['runAs', Session.getEffectiveUser().getEmail()],
                            ['startedAt', state.startedAt]]);
    return true;
  },

  domains: function () {
    const res = call_(() => AdminDirectory.Domains.list('my_customer'));
    const rows = [];
    (res.domains || []).forEach(d => {
      rows.push([d.domainName, 'domain', !!d.isPrimary, !!d.verified, '']);
      (d.domainAliases || []).forEach(a => rows.push([a.domainAliasName, 'alias', false, !!a.verified, d.domainName]));
    });
    appendRows_('domains', rows);
    return true;
  },

  orgunits: function () {
    const res = call_(() => AdminDirectory.Orgunits.list('my_customer', { type: 'all' }));
    const rows = [['/', '(root)', '', '', '']];
    (res.organizationUnits || []).forEach(o =>
      rows.push([o.orgUnitPath, o.name, o.description || '', o.parentOrgUnitPath || '', o.orgUnitId]));
    appendRows_('orgunits', rows);
    return true;
  },

  users: function (state, deadline) {
    return paginate_(state, deadline, token =>
      call_(() => AdminDirectory.Users.list({ customer: 'my_customer', maxResults: 500, pageToken: token,
                                             orderBy: 'email', projection: 'basic' })),
      res => (res.users || []).map(u => [
        u.primaryEmail, (u.name || {}).givenName || '', (u.name || {}).familyName || '', (u.name || {}).fullName || '',
        u.orgUnitPath || '/', !!u.suspended, !!u.archived, !!u.isAdmin, !!u.isDelegatedAdmin,
        (u.aliases || []).join(';'), (u.nonEditableAliases || []).join(';'),
        u.creationTime || '', u.lastLoginTime || '', !!u.isMailboxSetup, !!u.isEnrolledIn2Sv, u.customerId || '', u.id,
      ]), 'users');
  },

  /** Licence assignments per product. A product the customer doesn't have may error; that's recorded, never fatal. */
  licenses: function (state, deadline) {
    const cust = (readRaw_('users').find(u => u.customerId) || {}).customerId;
    if (!cust) { appendRows_('licenses', [['', '', '', '', '', 'No customer ID found on any user']]); return true; }
    return perItem_(state, deadline, LICENSE_PRODUCTS, (productId) => {
      const rows = [];
      let token;
      try {
        do {
          const res = call_(() => AdminLicenseManager.LicenseAssignments.listForProduct(productId, cust, { maxResults: 500, pageToken: token }));
          (res.items || []).forEach(a => rows.push([String(a.userId || '').toLowerCase(), a.productId || productId,
            a.productName || '', a.skuId || '', a.skuName || '', '']));
          token = res.nextPageToken;
        } while (token);
        return rows;
      } catch (e) {
        return [['', productId, '', '', '', e.message]];
      }
    }, 'licenses');
  },

  groups: function (state, deadline) {
    return paginate_(state, deadline, token =>
      call_(() => AdminDirectory.Groups.list({ customer: 'my_customer', maxResults: 200, pageToken: token })),
      res => (res.groups || []).map(g => [
        g.email, g.name || '', g.description || '', g.directMembersCount || 0,
        (g.aliases || []).join(';'), (g.nonEditableAliases || []).join(';'), !!g.adminCreated, g.id,
      ]), 'groups');
  },

  members: function (state, deadline) {
    const groups = readRawColumn_('groups', 'email');
    return perItem_(state, deadline, groups, (groupEmail) => {
      const rows = [];
      let token;
      do {
        const res = call_(() => AdminDirectory.Members.list(groupEmail, { maxResults: 200, pageToken: token }));
        (res.members || []).forEach(m => rows.push([groupEmail,
          m.email || (m.type === 'CUSTOMER' ? 'ALL_USERS_IN_ORG' : ''),
          m.role || '', m.type || '', m.status || '', m.delivery_settings || '', m.id || '']));
        token = res.nextPageToken;
      } while (token);
      return rows;
    }, 'members');
  },

  settings: function (state, deadline) {
    const groups = readRawColumn_('groups', 'email');
    const keys = RAW.settings.headers.slice(1, -2);
    return perItem_(state, deadline, groups, (groupEmail) => {
      try {
        const s = call_(() => AdminGroupsSettings.Groups.get(groupEmail));
        return [[groupEmail].concat(keys.map(k => s[k] === undefined ? '' : String(s[k])), ['', JSON.stringify(s)])];
      } catch (e) {
        return [[groupEmail].concat(keys.map(() => ''), [e.message, ''])];
      }
    }, 'settings');
  },

  drives: function (state, deadline) {
    return paginate_(state, deadline, token =>
      call_(() => Drive.Drives.list({ useDomainAdminAccess: true, pageSize: 100, pageToken: token,
        fields: 'nextPageToken,drives(id,name,createdTime,hidden,orgUnitId,restrictions)' })),
      res => (res.drives || []).map(d => {
        const r = d.restrictions || {};
        return [d.id, d.name, d.createdTime || '', !!d.hidden, d.orgUnitId || '', !!r.adminManagedRestrictions,
                !!r.copyRequiresWriterPermission, !!r.domainUsersOnly, !!r.driveMembersOnly,
                !!r.sharingFoldersRequiresOrganizerPermission];
      }), 'drives');
  },

  perms: function (state, deadline) {
    const ids = readRawColumn_('drives', 'driveId');
    const names = readRawColumn_('drives', 'name');
    return perItem_(state, deadline, ids, (driveId, idx) => {
      const rows = [];
      let token;
      try {
        do {
          const res = call_(() => Drive.Permissions.list(driveId, {
            supportsAllDrives: true, useDomainAdminAccess: true, pageSize: 100, pageToken: token,
            fields: 'nextPageToken,permissions(id,type,role,emailAddress,domain,displayName,deleted)' }));
          (res.permissions || []).forEach(p => rows.push([driveId, names[idx], p.type, p.role, p.emailAddress || '',
            p.domain || '', p.displayName || '', !!p.deleted, p.id, '']));
          token = res.nextPageToken;
        } while (token);
        return rows;
      } catch (e) {
        return [[driveId, names[idx], '', '', '', '', '', '', '', e.message]];
      }
    }, 'perms');
  },
};

function paginate_(state, deadline, fetchPage, toRows, key) {
  while (true) {
    if (Date.now() > deadline) return false;
    const res = fetchPage(state.cursor.pageToken || undefined);
    appendRows_(key, toRows(res));
    if (!res.nextPageToken) return true;
    state.cursor.pageToken = res.nextPageToken;
    saveState_(state);
  }
}

function perItem_(state, deadline, items, fetch, key) {
  let idx = state.cursor.idx || 0;
  let buffer = [];
  const flush = () => { if (buffer.length) { appendRows_(key, buffer); buffer = []; } state.cursor.idx = idx; saveState_(state); };
  while (idx < items.length) {
    if (Date.now() > deadline) { flush(); return false; }
    buffer = buffer.concat(fetch(items[idx], idx));
    idx++;
    if (buffer.length > 2000) flush();
  }
  flush();
  return true;
}

// ------------------------------------------------------------------------------------------ shared helpers

function call_(fn) {
  const transient = /rate limit|quota|backend error|internal error|429|500|502|503|timed out|too many|try again/i;
  for (let attempt = 0; ; attempt++) {
    try { return fn(); } catch (e) {
      if (attempt >= 6 || !transient.test(e.message)) throw e;
      Utilities.sleep(Math.min(32000, Math.pow(2, attempt) * 1000) + Math.floor(Math.random() * 1000));
    }
  }
}

function initRawSheet_(key) {
  const ss = SpreadsheetApp.getActive();
  const def = RAW[key];
  let sh = ss.getSheetByName(def.name);
  if (sh) sh.clear(); else sh = ss.insertSheet(def.name);
  ensureSize_(sh, 2, def.headers.length);
  sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/** getRange() throws outside the sheet's grid (new sheets are 1000 x 26), so grow it first. */
function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function appendRows_(key, rows) {
  if (!rows || !rows.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(RAW[key].name);
  ensureSize_(sh, sh.getLastRow() + rows.length, rows[0].length);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
}

function readRawColumn_(key, header) {
  return readRaw_(key).map(r => String(r[header]));
}

/** Reads a RAW tab into an array of objects keyed by header. */
function readRaw_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(RAW[key].name);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const h = vals[0].map(String);
  return vals.slice(1).map(r => { const o = {}; h.forEach((k, i) => { o[k] = String(r[i]).trim(); }); return o; });
}

function truthy_(v) { return ['true', '1', 'yes', 'y'].indexOf(String(v).trim().toLowerCase()) >= 0; }

function loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  return raw ? JSON.parse(raw) : null;
}

function saveState_(state) {
  state.updatedAt = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
}

function deleteContinueTriggers_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'continueRun')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function notify_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, 'Migration Mapping', 10); } catch (e) { /* trigger context */ }
  console.log(msg);
}

// ================================================================ Build.gs
/**
 * Build.gs — turns RAW_* tabs into a separate, client-facing review spreadsheet.
 * Mapping rule: every internal address user@domain -> user@<tempPrefix>domain; external addresses unchanged.
 */

const DRIVE_ROLE = { organizer: 'Manager', fileOrganizer: 'Content manager', writer: 'Contributor',
                     commenter: 'Commenter', reader: 'Viewer' };
const GROUP_ROLE = { OWNER: 'Owner', MANAGER: 'Manager', MEMBER: 'Member' };
const QUOTA_FLAG_PCT = 80;   // flag users at or above this % of their storage quota

// Review tab names — Export.gs reads these back, keep stable.
const T = { readme: 'Read Me & Sign-off', users: 'Google User Mapping', groups: 'Groups (Summary)',
            gmembers: 'GroupsDLs Mapping', memberships: 'User Memberships', drives: 'Shared Drives (Summary)',
            dacl: 'Shared Drive Mapping', ous: 'Org Units', aliases: 'Aliases', storage: 'Storage Summary', issues: 'Issues to Resolve' };

// [header, kind, width]  kind: g = generated (navy, warn on edit), c = client reviews/edits (yellow), i = internal (grey)
const COLS = {
  users: [['Migrate?', 'c', 80], ['Priority?', 'c', 70], ['First Name', 'g', 110], ['Last Name', 'g', 120],
    ['[Source] Original Google Account', 'g', 240], ['[Destination] Temp Google Account', 'g', 260],
    ['[Destination] Final Google Account', 'g', 240], ['OU Path', 'g', 300],
    ['[Source] Email address after cutover', 'g', 260], ['Email', 'c', 60], ['Contacts', 'c', 70],
    ['Calendar', 'c', 70], ['Drive', 'c', 60], ['Google Workspace licence', 'g', 220],
    ['Email Forwarding [Destination]', 'c', 180], ['Account status', 'g', 100], ['Last sign-in', 'g', 100],
    ['Gmail (GB)', 'g', 90], ['Drive (GB)', 'g', 90], ['Gmail + Drive (GB)', 'g', 100], ['Quota used % (all Google storage)', 'g', 130],
    ['Client Comments', 'c', 220], ['Status', 'i', 100], ['Cloudasta Comments', 'i', 220], ['Flags', 'g', 320]],
  groups: [['Migrate?', 'c', 80], ['Group Name', 'g', 200], ['[Source] Group email', 'g', 260],
    ['[Destination] Temp group email', 'g', 280], ['Description', 'g', 200], ['Group access type (derived)', 'g', 140],
    ['Do you need external users to send emails to the group?', 'c', 170], ['Direct members', 'g', 80],
    ['Aliases (see Aliases tab)', 'g', 110], ['Client Comments', 'c', 220], ['Flags', 'g', 300]],
  gmembers: [['Group Name', 'g', 200], ['Group email address', 'g', 260], ['[Destination] Temp group email', 'g', 280],
    ['Group access type', 'g', 130], ['Do you need external users to send emails to the group?', 'g', 160],
    ['Group member email address', 'g', 240], ['[Destination] Member email', 'g', 260], ['Member type', 'g', 130],
    ['Member role', 'c', 100], ['Member status', 'g', 100], ['Client Comments', 'c', 220]],
  memberships: [['User', 'g', 240], ['[Destination] User', 'g', 260], ['Group', 'g', 260],
    ['[Destination] Group', 'g', 280], ['Role', 'g', 90], ['Membership', 'g', 180]],
  drives: [['Migrate?', 'c', 80], ['[Source] Shared Drive Name', 'g', 260], ['[Source] Shared Drive ID', 'g', 200],
    ['[Destination] Shared Drive Name', 'c', 280], ['ACL entries', 'g', 80], ['Managers', 'g', 80],
    ['Created', 'g', 100], ['Restricted to domain users', 'g', 110], ['Client Comments', 'c', 220],
    ['[Destination] Shared Drive URL', 'i', 220], ['Status', 'i', 100], ['Flags', 'g', 320]],
  dacl: [['[Source] Shared Drive Name', 'g', 260], ['[Source] Shared Drive ID', 'g', 200],
    ['[Destination] Shared Drive Name', 'g', 280], ['[Source] Collaborator', 'g', 260],
    ['[Destination] Collaborator email address', 'g', 280], ['Collaborator type', 'g', 130], ['Role', 'c', 130],
    ['Client Comments', 'c', 220]],
  ous: [['OU Path', 'g', 420], ['Name', 'g', 200], ['Parent OU', 'g', 360], ['Description', 'g', 240],
    ['Users in OU', 'g', 90]],
  aliases: [['[Source] Target (user or group)', 'g', 260], ['[Source] Alias', 'g', 260], ['Target type', 'g', 90],
    ['Alias type', 'g', 170], ['[Destination] Target', 'g', 280], ['[Destination] Alias', 'g', 280],
    ['Migrate alias?', 'c', 110], ['Client Comments', 'c', 220], ['Flags', 'g', 320]],
  storage: [['Metric', 'g', 380], ['Value', 'g', 200], ['Notes', 'g', 460]],
  issues: [['Severity', 'g', 80], ['Area', 'g', 110], ['Object', 'g', 300], ['Issue', 'g', 420],
    ['Recommended action', 'g', 420]],
};

const GROUP_SETTING_COLS = ['whoCanJoin', 'whoCanViewMembership', 'whoCanViewGroup', 'whoCanPostMessage',
  'allowExternalMembers', 'whoCanContactOwner', 'whoCanDiscoverGroup', 'whoCanLeaveGroup', 'messageModerationLevel',
  'whoCanModerateMembers', 'whoCanModerateContent', 'allowWebPosting', 'isArchived', 'membersCanPostAsTheGroup',
  'includeInGlobalAddressList', 'replyTo', 'customReplyTo', 'enableCollaborativeInbox'];

function menuBuild() {
  const ui = SpreadsheetApp.getUi();
  const st = loadState_();
  if (!st || st.status !== 'COMPLETE') {
    ui.alert('Run "1. Collect source data" first and wait until the Status setting says it is complete.');
    return;
  }
  const cfg = getSettings_();
  notify_('Building review spreadsheet…');
  const model = buildModel_(cfg);

  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty(REVIEW_ID_KEY);
  if (oldId) {
    try { const old = SpreadsheetApp.openById(oldId); if (old.getName().indexOf('[SUPERSEDED]') !== 0) old.rename('[SUPERSEDED] ' + old.getName()); }
    catch (e) { /* deleted */ }
  }

  const stamp = Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyy-MM-dd');
  const ss = SpreadsheetApp.create(`${cfg.project} — Migration Mappings for Review (${stamp})`);
  writeReadme_(ss.getSheets()[0], cfg, model);
  writeTab_(ss, T.users, COLS.users, model.users, { yesno: ['Migrate?'], checkbox: ['Priority?', 'Email', 'Contacts', 'Calendar', 'Drive'],
            number: { 'Gmail (GB)': '#,##0.00', 'Drive (GB)': '#,##0.00', 'Gmail + Drive (GB)': '#,##0.00',
                      'Quota used % (all Google storage)': '0"%"' } });
  writeTab_(ss, T.aliases, COLS.aliases, model.aliases, { list: { 'Migrate alias?': ['Yes', 'No', 'N/A (automatic)'] } });
  writeTab_(ss, T.groups, COLS.groups.concat(GROUP_SETTING_COLS.map(k => [k, 'g', 150])), model.groups,
            { yesno: ['Migrate?'], list: { 'Do you need external users to send emails to the group?': ['Yes', 'No'] } });
  writeTab_(ss, T.gmembers, COLS.gmembers, model.gmembers, { list: { 'Member role': ['Owner', 'Manager', 'Member'] } });
  writeTab_(ss, T.memberships, COLS.memberships, model.memberships, {});
  writeTab_(ss, T.drives, COLS.drives, model.drives, { yesno: ['Migrate?'] });
  writeTab_(ss, T.dacl, COLS.dacl, model.dacl, { list: { Role: Object.keys(DRIVE_ROLE).map(k => DRIVE_ROLE[k]) } });
  writeTab_(ss, T.ous, COLS.ous, model.ous, {});
  writeTab_(ss, T.storage, COLS.storage, model.storage, {});
  writeTab_(ss, T.issues, COLS.issues, model.issues, {});
  ss.setActiveSheet(ss.getSheets()[0]);

  props.setProperty(REVIEW_ID_KEY, ss.getId());
  setSetting_('reviewUrl', ss.getUrl());
  const high = model.issues.filter(r => r[0] === 'High').length;
  setSetting_('status', `Review built ${stamp}: ${model.users.length} users, ${model.groups.length} groups, ` +
              `${model.drives.length} shared drives, ${model.issues.length} issues (${high} high). Next: send for review.`);
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial;font-size:14px;line-height:1.5">
       <p><b>Review spreadsheet created.</b></p>
       <p><a href="${ss.getUrl()}" target="_blank">Open the review spreadsheet</a></p>
       <p>${model.users.length} users · ${model.groups.length} groups · ${model.drives.length} shared drives<br>
          ${model.issues.length} issues to resolve (${high} high) — see the <i>Issues to Resolve</i> tab.</p>
       <p>Share it with the client's approver (Share button), or File → Download → .xlsx and email it.</p></div>`)
    .setWidth(460).setHeight(260);
  ui.showModalDialog(html, 'Migration Mapping');
}

// ------------------------------------------------------------------------------------------ model

function buildModel_(cfg) {
  const R = {};
  Object.keys(RAW).forEach(k => { if (k !== 'runinfo') R[k] = readRaw_(k); });
  const srcDomains = new Set(R.domains.map(d => d.domainName.toLowerCase()));
  const users = {}, groups = {}, aliasTo = {};
  R.users.forEach(u => { users[u.primaryEmail.toLowerCase()] = u; });
  R.groups.forEach(g => { groups[g.email.toLowerCase()] = g; });
  R.users.forEach(u => splitList_(u.aliases + ';' + u.nonEditableAliases).forEach(a => { aliasTo[a] = u.primaryEmail.toLowerCase(); }));
  R.groups.forEach(g => splitList_(g.aliases + ';' + g.nonEditableAliases).forEach(a => { aliasTo[a] = g.email.toLowerCase(); }));

  // ---- licences (Enterprise License Manager; optional)
  const lic = {};
  R.licenses.filter(l => l.email && !l.error).forEach(l => { (lic[l.email] = lic[l.email] || []).push(l); });
  const wsErr = R.licenses.find(l => l.productId === 'Google-Apps' && l.error);
  const licOk = !wsErr && R.licenses.some(l => l.email);
  const licenceOf = e => {
    if (!licOk) return '';
    const all = lic[e] || [];
    const pick = ids => all.filter(l => ids.indexOf(l.productId) >= 0).map(l => l.skuName || l.skuId);
    const ws = pick(['Google-Apps']);
    if (ws.length) return ws.join('; ');
    const other = pick(['101034']).concat(pick(['101001', '101005']));
    return other.length ? other.join('; ') : 'No licence';
  };

  const isInternal = e => e.indexOf('@') > 0 && srcDomains.has(e.split('@').pop().toLowerCase());
  const canon = e => { const x = String(e).trim().toLowerCase(); return aliasTo[x] || x; };
  const temp = e => { const x = canon(e); if (!isInternal(x)) return x; const p = x.split('@'); return `${p[0]}@${cfg.tempPrefix}${p[1]}`; };
  const tempRaw = e => { const x = String(e).trim().toLowerCase(); if (!isInternal(x)) return x; const p = x.split('@'); return `${p[0]}@${cfg.tempPrefix}${p[1]}`; };
  const kind = e => { const x = canon(e); return users[x] ? 'User' : groups[x] ? 'Group' : isInternal(x) ? 'Internal (unknown/deleted)' : 'External'; };

  const issues = [];
  const issue = (sev, area, obj, text, action) => issues.push([sev, area, obj, text, action]);

  // ---- usage (Admin SDK Reports; optional)
  const tenant = {};
  R.tenantusage.forEach(r => { tenant[r.metric] = r.value; });
  const usage = {};
  R.userusage.forEach(r => {
    // "What we migrate" = Gmail + Drive. Google's own total also counts Photos, which we don't migrate.
    const g = Number(r.gmailMB), d = Number(r.driveMB);
    r.migMB = isNaN(g) && isNaN(d) ? '' : (isNaN(g) ? 0 : g) + (isNaN(d) ? 0 : d);
    usage[r.email.toLowerCase()] = r;
  });
  const usageOk = tenant.status === 'OK';
  const gb = mb => (mb === '' || mb === undefined || isNaN(Number(mb))) ? '' : Math.round(Number(mb) / 1024 * 100) / 100;
  if (!usageOk) issue('Info', 'Storage', '-', `Storage usage unavailable from the Reports API${tenant.detail ? ': ' + tenant.detail.slice(0, 300) : ''}`,
                      'Usage columns are blank. Check Admin console → Reports, or re-run collection in a day or two.');

  // ---- users
  const userRows = R.users.slice().sort(byKey_('primaryEmail')).map(u => {
    const e = u.primaryEmail.toLowerCase();
    let status = 'Active';
    if (truthy_(u.suspended)) status = 'Suspended';
    if (truthy_(u.archived)) status = 'Archived';
    const last = daysSince_(u.lastLoginTime);
    const flags = [];
    if (status !== 'Active') flags.push(status.toLowerCase());
    if (last === null) flags.push('never signed in'); else if (last > cfg.staleDays) flags.push(`no sign-in for ${last} days`);
    if (truthy_(u.isAdmin)) flags.push('super admin');
    if (!u.givenName || !u.familyName) {
      flags.push('missing first/last name');
      issue('High', 'Users', e, 'Missing first or last name', 'Names are required to create the destination account.');
    }
    const licence = licenceOf(e);
    if (licence === 'No licence' && status === 'Active') {
      flags.push('no licence assigned');
      issue('Low', 'Users', e, 'Active account with no Google Workspace licence', 'Confirm whether it needs a licence at the destination.');
    }
    if (/^Cloud Identity/.test(licence)) flags.push('Cloud Identity only (no Workspace plan)');
    const us = usage[e];
    const pct = us && us.usedPct !== '' ? Number(us.usedPct) : '';
    if (pct !== '' && pct >= QUOTA_FLAG_PCT) {
      flags.push(`${pct}% of storage quota used`);
      issue('Low', 'Users', e, `Using ${pct}% of storage quota (Gmail + Drive: ${gb(us.migMB)} GB)`,
            'Make sure the destination licence has enough storage for this account.');
    }
    if (usageOk && !us) flags.push('no usage data (new or recently changed account?)');
    const migrate = status === 'Active' && last !== null ? 'Yes' : 'Review';
    if (migrate === 'Review') issue('Medium', 'Users', e, flags.join('; '), 'Client to confirm Migrate? = Yes or No.');
    return [migrate, false, u.givenName, u.familyName, e, temp(e), e, u.orgUnitPath || '/',
      e.split('@')[0] + '@' + cfg.archivePrefix + e.split('@')[1], true, true, true, true,
      licence, '', status, last === null ? 'Never' : u.lastLoginTime.slice(0, 10),
      us ? gb(us.gmailMB) : '', us ? gb(us.driveMB) : '', us ? gb(us.migMB) : '', pct,
      '', '', '', flags.join('; ')];
  });

  // ---- groups
  const settings = {};
  R.settings.forEach(s => { settings[s.groupEmail.toLowerCase()] = s; });
  const membersBy = {};
  R.members.forEach(m => { (membersBy[m.groupEmail.toLowerCase()] = membersBy[m.groupEmail.toLowerCase()] || []).push(m); });
  const groupRows = [], gmemberRows = [], directOf = {};
  R.groups.slice().sort(byKey_('email')).forEach(g => {
    const ge = g.email.toLowerCase(), gt = temp(ge), s = settings[ge] || {};
    const acc = accessType_(s);
    const ext = !s.whoCanPostMessage ? 'Unknown' : s.whoCanPostMessage === 'ANYONE_CAN_POST' ? 'Yes' : 'No';
    const mbs = membersBy[ge] || [];
    const flags = [];
    const nExt = mbs.filter(m => m.memberEmail && m.memberEmail !== 'ALL_USERS_IN_ORG' && !isInternal(m.memberEmail)).length;
    if (!mbs.length) flags.push('no members');
    if (nExt) flags.push(`${nExt} external member(s)`);
    if (mbs.some(m => m.type === 'GROUP')) flags.push('contains nested group(s)');
    if (mbs.some(m => m.memberEmail === 'ALL_USERS_IN_ORG')) {
      flags.push("'all users in organization' member");
      issue('Medium', 'Groups', ge, "Has 'all users in organization' as a member", 'Re-add at the destination after cutover; it cannot be mapped to a placeholder address.');
    }
    if (s.error) { flags.push('settings unreadable'); issue('Medium', 'Groups', ge, `Settings could not be read: ${s.error}`, 'Capture settings manually.'); }
    if (acc === 'Custom') flags.push('custom access settings — see settings columns');
    groupRows.push(['Yes', g.name, ge, gt, g.description, acc, ext, mbs.length,
      splitList_(g.aliases).length + splitList_(g.nonEditableAliases).length, '', flags.join('; ')]
      .concat(GROUP_SETTING_COLS.map(k => s[k] || '')));
    const order = { OWNER: 0, MANAGER: 1, MEMBER: 2 };
    mbs.slice().sort((a, b) => (order[a.role] - order[b.role]) || a.memberEmail.localeCompare(b.memberEmail)).forEach(m => {
      const all = m.memberEmail === 'ALL_USERS_IN_ORG';
      const me = all ? 'ALL_USERS_IN_ORG' : m.memberEmail.toLowerCase();
      let mtype = { USER: 'User', GROUP: 'Group', CUSTOMER: 'All users in org' }[m.type] || m.type;
      if (!all && !isInternal(me)) mtype += ' (external)';
      gmemberRows.push([g.name, ge, gt, acc, ext, me, all ? 'Re-add after cutover' : temp(me), mtype,
                        GROUP_ROLE[m.role] || m.role, m.status, '']);
      if (!all) (directOf[canon(me)] = directOf[canon(me)] || []).push([ge, GROUP_ROLE[m.role] || m.role]);
    });
  });

  // ---- per-user memberships, including those inherited through nested groups
  const expand = (member, seen) => {
    let out = [];
    (directOf[member] || []).forEach(([grp]) => { if (!seen.has(grp)) { seen.add(grp); out.push(grp); out = out.concat(expand(grp, seen)); } });
    return out;
  };
  const membershipRows = [];
  Object.keys(users).sort().forEach(e => {
    const direct = (directOf[e] || []).slice().sort((a, b) => a[0].localeCompare(b[0]));
    const dset = new Set(direct.map(d => d[0]));
    direct.forEach(([grp, role]) => membershipRows.push([e, temp(e), grp, temp(grp), role, 'Direct']));
    Array.from(new Set(expand(e, new Set()))).filter(g => !dset.has(g)).sort()
      .forEach(grp => membershipRows.push([e, temp(e), grp, temp(grp), 'Member', 'Inherited (nested group)']));
  });

  // ---- shared drives
  const permsBy = {};
  R.perms.forEach(p => { (permsBy[p.driveId] = permsBy[p.driveId] || []).push(p); });
  const roleOrder = Object.keys(DRIVE_ROLE);
  const driveRows = [], daclRows = [];
  R.drives.slice().sort(byKey_('name')).forEach(d => {
    const all = permsBy[d.driveId] || [];
    const ps = all.filter(p => !p.error), errs = all.filter(p => p.error);
    const label = `${d.name} (${d.driveId})`;
    const dname = cfg.drivePrefix + d.name;
    const flags = [];
    if (errs.length) { flags.push('ACL unreadable'); issue('High', 'Shared Drives', label, `Permissions could not be read: ${errs[0].error}`, 'Check manually in the Admin console.'); }
    const nMgr = ps.filter(p => p.role === 'organizer').length;
    if (!nMgr) { flags.push('no Manager (orphaned)'); issue('High', 'Shared Drives', label, 'Shared drive has no Manager', 'Client to nominate a Manager for the destination drive.'); }
    const nExt = ps.filter(p => (p.type === 'user' || p.type === 'group') && p.emailAddress && !isInternal(p.emailAddress)).length;
    if (nExt) flags.push(`${nExt} external collaborator(s)`);
    const wide = ps.filter(p => p.type === 'domain' || p.type === 'anyone');
    if (wide.length) {
      flags.push('domain/anyone permission');
      issue('Medium', 'Shared Drives', label, `Drive-level '${wide[0].type}' permission (${wide[0].domain || 'anyone'})`,
            'Decide the destination target. It is not included in the destination ACL file.');
    }
    if (ps.some(p => truthy_(p.deleted))) flags.push('deleted account in ACL');
    if (truthy_(d.hidden)) flags.push('hidden drive');
    driveRows.push(['Yes', d.name, d.driveId, dname, ps.length, nMgr, d.createdTime.slice(0, 10),
                    truthy_(d.domainUsersOnly) ? 'Yes' : 'No', '', '', '', flags.join('; ')]);
    ps.slice().sort((a, b) => (roleOrder.indexOf(a.role) - roleOrder.indexOf(b.role)) || a.emailAddress.localeCompare(b.emailAddress))
      .forEach(p => {
        let src, dst, ctype;
        if (p.type === 'domain' || p.type === 'anyone') {
          src = `${p.type}:${p.domain || 'anyone'}`; dst = 'Decision needed'; ctype = p.type === 'domain' ? 'Domain' : 'Anyone';
        } else {
          src = p.emailAddress.toLowerCase(); dst = temp(src);
          ctype = p.type === 'user' ? kind(src) : (isInternal(src) ? 'Group' : 'Group (external)');
          if (ctype.indexOf('Internal (unknown') === 0)
            issue('Low', 'Shared Drives', d.name, `ACL references ${src}, which is not a current user or group`, 'Usually a deleted account. It will be skipped at export.');
        }
        daclRows.push([d.name, d.driveId, dname, src, dst, ctype, DRIVE_ROLE[p.role] || p.role, '']);
      });
  });

  // ---- org units
  const ouCount = {};
  R.users.forEach(u => { ouCount[u.orgUnitPath || '/'] = (ouCount[u.orgUnitPath || '/'] || 0) + 1; });
  const ouRows = R.orgunits.slice().sort(byKey_('orgUnitPath'))
    .map(o => [o.orgUnitPath, o.name, o.parentOrgUnitPath, o.description, ouCount[o.orgUnitPath] || 0]);

  // ---- aliases: one row per alias (users first, then groups), target in the first column
  const aliasRows = [];
  const addAliases = (target, type, editable, automatic) => {
    const t = target.toLowerCase();
    splitList_(editable).sort().forEach(a => aliasRows.push([t, a, type, 'User-created', temp(t), tempRaw(a), 'Yes', '',
      isInternal(a) ? '' : 'alias domain is not a domain of this tenant']));
    splitList_(automatic).sort().forEach(a => aliasRows.push([t, a, type, 'Automatic (domain alias)', temp(t), '',
      'N/A (automatic)', '', `created automatically from the domain alias ${a.split('@')[1]}; not exported`]));
  };
  R.users.slice().sort(byKey_('primaryEmail')).forEach(u => addAliases(u.primaryEmail, 'User', u.aliases, u.nonEditableAliases));
  R.groups.slice().sort(byKey_('email')).forEach(g => addAliases(g.email, 'Group', g.aliases, g.nonEditableAliases));
  const nAuto = aliasRows.filter(r => r[3] !== 'User-created').length;
  if (nAuto) issue('Info', 'Aliases', '-', `${nAuto} automatic alias(es) come from domain aliases`,
                   'They are recreated by Google when the domain alias exists at the destination, so they are not exported.');
  const primaryDomains = R.domains.filter(d => d.type !== 'alias').map(d => d.domainName.toLowerCase());
  if (primaryDomains.length > 1) issue('Info', 'Tenant', primaryDomains.join(', '), 'The source tenant has more than one domain',
    `The destination needs every placeholder domain added and verified: ${primaryDomains.map(d => cfg.tempPrefix + d).join(', ')}.`);

  // ---- licence summary (Read Me)
  const licCount = {};
  userRows.forEach(r => { if (r[13]) licCount[r[13]] = (licCount[r[13]] || 0) + 1; });
  const licSummary = Object.keys(licCount).sort((a, b) => licCount[b] - licCount[a]).map(k => `${k} (${licCount[k]})`).join(', ');
  if (!licOk) issue('Info', 'Licences', '-', `Licence data unavailable${wsErr ? ': ' + wsErr.error.slice(0, 300) : ''}`,
                    'The licence column is blank. Check licences in Admin console → Billing → Licence settings.');

  // ---- storage summary (tenant totals from the Reports API + sums of the per-user figures)
  const fmtGB = mb => { const v = gb(mb); return v === '' ? 'n/a' : `${v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} GB`; };
  const sumMB = list => list.reduce((t, r) => t + (Number((usage[r[4]] || {}).migMB) || 0), 0);
  const storageRows = [];
  const tenantMigMB = ['gmail_used_quota_in_mb', 'drive_used_quota_in_mb', 'team_drive_used_quota_in_mb']
    .reduce((t, k) => t + (Number(tenant[k]) || 0), 0);
  if (usageOk) {
    const quota = Number(tenant.total_quota_in_mb);
    storageRows.push(
      ['Report date (Pacific time)', tenant.reportDate, 'Google storage reports lag 1–3 days behind today.'],
      ['TENANT TOTALS', '', ''],
      ['Total to migrate', fmtGB(tenantMigMB), 'Gmail + My Drive + shared drives (Google Photos is not migrated and is excluded)'],
      ['Gmail', fmtGB(tenant.gmail_used_quota_in_mb), ''],
      ['Drive (My Drive)', fmtGB(tenant.drive_used_quota_in_mb), ''],
      ['Shared drives (all)', fmtGB(tenant.team_drive_used_quota_in_mb), 'Google provides no per-drive figure; see the CloudM scan for per-drive sizes'],
      ['Total storage quota', quota > 0 ? fmtGB(quota) : 'Not reported', quota > 0 ? '' : 'Google reports no fixed quota for some editions'],
      ['Users (per report)', String(tenant.num_users || ''), ''],
      [`Users using ${QUOTA_FLAG_PCT}% or more of quota`, String(tenant.num_users_used_quota_ge_80percent || '0'), ''],
      ['Users using 50–80% of quota', String(tenant.num_users_used_quota_ge_50_lt_80percent || '0'), ''],
      ['', '', ''],
      ['USER ACCOUNTS: Gmail + Drive (sum of per-user figures)', '', ''],
      ['All user accounts', fmtGB(sumMB(userRows)),
        `${userRows.length} accounts (${userRows.filter(r => usage[r[4]]).length} with usage data)`],
      ['Default Migrate? = Yes', fmtGB(sumMB(userRows.filter(r => r[0] === 'Yes'))), `${userRows.filter(r => r[0] === 'Yes').length} accounts`],
      ['Default Migrate? = Review', fmtGB(sumMB(userRows.filter(r => r[0] === 'Review'))), 'Suspended, archived or never-signed-in accounts awaiting a decision'],
      ['', '', ''],
      ['LARGEST ACCOUNTS (Gmail + Drive)', '', '']);
    userRows.filter(r => usage[r[4]]).sort((a, b) => (Number(usage[b[4]].migMB) || 0) - (Number(usage[a[4]].migMB) || 0))
      .slice(0, 10).forEach((r, i) => storageRows.push([`${i + 1}. ${r[4]}`, fmtGB(usage[r[4]].migMB),
        `Gmail ${fmtGB(usage[r[4]].gmailMB)} · Drive ${fmtGB(usage[r[4]].driveMB)} · ${r[15]}`]));
  } else {
    storageRows.push(['Storage usage', 'Unavailable', tenant.detail || 'The Reports API returned no data for the last 7 days.']);
  }

  const sev = { High: 0, Medium: 1, Low: 2, Info: 3 };
  issues.sort((a, b) => (sev[a[0]] - sev[b[0]]) || a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]));
  return { users: userRows, groups: groupRows, gmembers: gmemberRows, memberships: membershipRows,
           drives: driveRows, dacl: daclRows, ous: ouRows, storage: storageRows, aliases: aliasRows, issues: issues, tempRaw: tempRaw,
           tenantUsage: usageOk ? tenant : null, tenantMigMB: tenantMigMB, fmtGB: fmtGB, licSummary: licSummary };
}

/** Best-effort mapping of Groups Settings to Admin console "access type" presets; anything else = Custom. */
function accessType_(s) {
  if (!s || s.error || !s.whoCanPostMessage) return 'Unknown';
  const post = s.whoCanPostMessage, view = s.whoCanViewGroup, join = s.whoCanJoin;
  if (['ALL_MANAGERS_CAN_POST', 'ALL_OWNERS_CAN_POST', 'NONE_CAN_POST'].indexOf(post) >= 0) return 'Announcement only';
  if (post === 'ANYONE_CAN_POST' && ['ALL_IN_DOMAIN_CAN_JOIN', 'ANYONE_CAN_JOIN'].indexOf(join) >= 0) return 'Public';
  if (post === 'ALL_IN_DOMAIN_CAN_POST') return 'Team';
  if (['ALL_MEMBERS_CAN_POST', 'ANYONE_CAN_POST'].indexOf(post) >= 0 && view === 'ALL_MEMBERS_CAN_VIEW' &&
      ['INVITED_CAN_JOIN', 'CAN_REQUEST_TO_JOIN'].indexOf(join) >= 0) return 'Restricted';
  return 'Custom';
}

// ------------------------------------------------------------------------------------------ writing

function writeTab_(ss, name, cols, rows, opt) {
  const sh = ss.insertSheet(name);
  const n = cols.length, headers = cols.map(c => c[0]);
  ensureSize_(sh, Math.max(rows.length, 1) + 1 + 200, n);   // +200 spare rows for data validation
  const fill = { g: '#1F3864', c: '#FFE699', i: '#BFBFBF' };
  sh.getRange(1, 1, 1, n).setValues([headers]).setFontWeight('bold').setWrap(true).setVerticalAlignment('top')
    .setBackgrounds([cols.map(c => fill[c[1]])]).setFontColors([cols.map(c => c[1] === 'g' ? '#FFFFFF' : '#000000')]);
  sh.setRowHeight(1, 48);
  cols.forEach((c, i) => sh.setColumnWidth(i + 1, c[2]));
  sh.setFrozenRows(1);
  const last = Math.max(rows.length, 1) + 1;
  if (rows.length) {
    const checkbox = new Set((opt.checkbox || []).map(h => headers.indexOf(h)));
    const numFmt = {};
    Object.keys(opt.number || {}).forEach(h => { numFmt[headers.indexOf(h)] = opt.number[h]; });
    // Text format everywhere except checkbox/number columns, so IDs/emails are never coerced.
    headers.forEach((h, i) => { if (!checkbox.has(i)) sh.getRange(2, i + 1, rows.length, 1).setNumberFormat(i in numFmt ? numFmt[i] : '@'); });
    sh.getRange(2, 1, rows.length, n).setValues(rows.map(r => r.map((v, i) => {
      if (checkbox.has(i)) return !!v;
      if (i in numFmt) return v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v);
      return v === null || v === undefined ? '' : String(v);
    })));
    checkbox.forEach(i => sh.getRange(2, i + 1, rows.length, 1).insertCheckboxes());
    const fi = headers.indexOf('Flags');
    if (fi >= 0) sh.getRange(2, fi + 1, rows.length, 1).setBackgrounds(rows.map(r => [r[fi] ? '#F8CBAD' : null]));
  }
  (opt.yesno || []).forEach(h => setList_(sh, headers.indexOf(h) + 1, last, ['Yes', 'No', 'Review']));
  Object.keys(opt.list || {}).forEach(h => setList_(sh, headers.indexOf(h) + 1, last, opt.list[h]));
  sh.getRange(1, 1, last, n).createFilter();
  // Warn (not block) when someone edits a generated column.
  cols.forEach((c, i) => { if (c[1] === 'g') sh.getRange(2, i + 1, last - 1, 1).protect().setWarningOnly(true).setDescription('Generated from source tenant'); });
  if (name === T.issues && rows.length) {
    const colors = { High: '#F4B6B6', Medium: '#FCE4B6', Low: '#E2EFDA', Info: '#DDEBF7' };
    sh.getRange(2, 1, rows.length, 1).setBackgrounds(rows.map(r => [colors[r[0]] || null]));
  }
}

function setList_(sh, col, last, values) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  sh.getRange(2, col, last - 1 + 200, 1).setDataValidation(rule);
}

function writeReadme_(sh, cfg, m) {
  sh.setName(T.readme);
  const review = m.users.filter(u => u[0] === 'Review').length;
  const high = m.issues.filter(i => i[0] === 'High').length;
  const lines = [
    [`${cfg.project} — Google Workspace migration mappings`, 'title'],
    ['', ''],
    ['Purpose', 'h'],
    ['We generated these mappings directly from your Google Workspace tenant. Your new environment will be built as a ' +
     'copy of your current one: the same users, groups, group memberships and roles, shared drives and their sharing. ' +
     'Please review each tab, correct anything in the yellow columns, and sign off below.', ''],
    ['', ''],
    ['How to review', 'h'],
    ['• Navy columns come from your tenant and should not be edited. Yellow columns are for you to review or edit. Grey columns are for Cloudasta.', ''],
    ["• 'Migrate?': choose Yes or No for every row. Rows marked 'Review' need your decision (for example suspended or inactive accounts).", ''],
    [`• Until cutover, new accounts and groups use a temporary address (name@${cfg.tempPrefix}yourdomain). At cutover they are renamed to the final address.`, ''],
    ["• 'Issues to Resolve' lists the items that need a decision before we can continue.", ''],
    ['• Temporary passwords are not in this spreadsheet. They are generated securely when the migration is configured.', ''],
    ['', ''],
    ['Summary', 'h'],
    [`Users: ${m.users.length} (${review} need a decision)`, ''],
    [`Groups: ${m.groups.length} (${m.gmembers.length} memberships)`, ''],
    [`Shared drives: ${m.drives.length} (${m.dacl.length} sharing entries)`, ''],
    [`Org units: ${m.ous.length}`, ''],
    [`Aliases: ${m.aliases.length} (see the Aliases tab)`, ''],
    [`Licences: ${m.licSummary || 'not available'}`, ''],
    [m.tenantUsage ? `Storage to migrate: ${m.fmtGB(m.tenantMigMB)} in total (Gmail ${m.fmtGB(m.tenantUsage.gmail_used_quota_in_mb)}, ` +
      `Drive ${m.fmtGB(m.tenantUsage.drive_used_quota_in_mb)}, shared drives ${m.fmtGB(m.tenantUsage.team_drive_used_quota_in_mb)}) ` +
      `as of ${m.tenantUsage.reportDate}. See the Storage Summary tab.` : 'Storage used: not available (see Issues to Resolve)', ''],
    [`Issues: ${m.issues.length} (${high} high priority)`, ''],
    ['', ''],
    ['Sign-off', 'h'],
    ['Approved by (name):', 'sign'], ['Title:', 'sign'], ['Date:', 'sign'],
    ['By signing, I confirm the user, group and shared drive mappings in this spreadsheet are correct.', ''],
    ['', ''],
    [`Generated ${new Date().toISOString().slice(0, 10)} · tool v${TOOL_VERSION}`, 'small'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines.map(l => [l[0]])).setWrap(true);
  sh.setColumnWidth(1, 820); sh.setColumnWidth(2, 320);
  lines.forEach((l, i) => {
    const r = sh.getRange(i + 1, 1);
    if (l[1] === 'title') r.setFontSize(16).setFontWeight('bold');
    if (l[1] === 'h') r.setFontWeight('bold').setFontSize(12);
    if (l[1] === 'small') r.setFontColor('#808080').setFontSize(9);
    if (l[1] === 'sign') sh.getRange(i + 1, 2).setBackground('#FFE699').setBorder(true, true, true, true, false, false);
  });
}

// ------------------------------------------------------------------------------------------ utils

function splitList_(s) { return String(s || '').split(';').map(x => x.trim().toLowerCase()).filter(Boolean); }
function byKey_(k) { return (a, b) => String(a[k]).toLowerCase().localeCompare(String(b[k]).toLowerCase()); }
function daysSince_(iso) {
  if (!iso || iso.indexOf('1970') === 0) return null;
  const t = Date.parse(iso);
  return isNaN(t) ? null : Math.floor((Date.now() - t) / 86400000);
}

// ================================================================ Export.gs
/**
 * Export.gs — reads the APPROVED review spreadsheet and produces one .zip in Drive containing:
 *   CloudM Migrate:   cloudm_users.csv, cloudm_shared_drives.csv, cloudm_address_replacements.csv
 *   Destination build (1:1 rebuild of source on placeholder addresses):
 *                     dest_org_units.csv, dest_users.csv, dest_groups.csv, dest_group_members.csv,
 *                     dest_shared_drives.csv, dest_shared_drive_acl.csv
 *   README.txt        what each file is + every warning raised during export
 * CloudM formats follow CloudM's "Bulk Add/Import Items" and "Address Replacements" articles.
 * Created by Luis Melgar - Cloud Migration Engineer at Cloudasta.
 */

const EXPORT_IDS_KEY = 'G2G_EXPORT_IDS';
const ROLE_FROM_LABEL = { 'Manager': 'organizer', 'Content manager': 'fileOrganizer', 'Contributor': 'writer',
                          'Commenter': 'commenter', 'Viewer': 'reader' };

function menuExport() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getSettings_();
  if (!cfg.reviewUrl) { ui.alert('Run "2. Build review spreadsheet" first.'); return; }
  let rv;
  try { rv = SpreadsheetApp.openByUrl(cfg.reviewUrl); }
  catch (e) { ui.alert(`Cannot open the review spreadsheet in Settings → "Review spreadsheet URL":\n${e.message}`); return; }

  const users = readTab_(rv, T.users), groups = readTab_(rv, T.groups), gmembers = readTab_(rv, T.gmembers);
  const drives = readTab_(rv, T.drives), dacl = readTab_(rv, T.dacl), aliases = readTab_(rv, T.aliases);
  if (!users || !groups || !gmembers || !drives || !dacl || !aliases) {
    ui.alert('The review spreadsheet is missing one of its tabs (for example "Aliases", added in v2.2.0). Point Settings → ' +
             '"Review spreadsheet URL" at the approved review spreadsheet. If it was built with an older version of the tool, run step 2 again.');
    return;
  }

  // ---- gate: every Migrate? decided, sign-off present
  const undecided = [].concat(
    users.filter(r => !isYesNo_(r['Migrate?'])).map(r => r['[Source] Original Google Account']),
    groups.filter(r => !isYesNo_(r['Migrate?'])).map(r => r['[Source] Group email']),
    drives.filter(r => !isYesNo_(r['Migrate?'])).map(r => r['[Source] Shared Drive Name']));
  if (undecided.length) {
    const a = ui.alert(`${undecided.length} row(s) still have Migrate? = Review or blank, for example:\n` +
      undecided.slice(0, 8).join('\n') + '\n\nTreat them as "No" and continue?', ui.ButtonSet.YES_NO);
    if (a !== ui.Button.YES) return;
  }
  const signer = signOffValue_(rv, 'Approved by (name):');
  if (!signer) {
    if (ui.alert('The client sign-off (Approved by) is empty. Export anyway?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }

  const warn = [];
  const yes = r => String(r['Migrate?']).trim() === 'Yes';
  const incUsers = users.filter(yes), incGroups = groups.filter(yes), incDrives = drives.filter(yes);
  const lc = s => String(s || '').trim().toLowerCase();

  // Addresses the review knows about (source -> placeholder), and which of them are being migrated.
  const allTemp = new Set(users.map(u => lc(u['[Destination] Temp Google Account']))
    .concat(groups.map(g => lc(g['[Destination] Temp group email']))));
  const incTemp = new Set(incUsers.map(u => lc(u['[Destination] Temp Google Account']))
    .concat(incGroups.map(g => lc(g['[Destination] Temp group email']))));
  const tempDomains = new Set(Array.from(allTemp).map(a => a.split('@')[1]));
  const isPlaceholder = a => tempDomains.has(String(a).split('@')[1]);

  // ---- passwords (same value in CloudM and destination user files)
  const pw = {};
  incUsers.forEach(u => { pw[lc(u['[Destination] Temp Google Account'])] = genPassword_(cfg.pwLength); });
  incUsers.forEach(u => { if (!u['First Name'] || !u['Last Name']) warn.push(`User ${u['[Source] Original Google Account']} has no first/last name, which is required to create the account.`); });

  const files = {};

  // 1. CloudM users
  files['cloudm_users.csv'] = toCsv_(
    ['Migrate', 'Priority', 'ExportObjectType', 'ImportObjectType', 'ExportName', 'ImportName', 'GivenName', 'FamilyName',
     'Password', 'ArchivePath', 'DocumentsPath', 'MigrateMail', 'MigrateContacts', 'MigrateCalendar', 'MigrateDrive',
     'MigrateSites', 'MigrateTasks', 'ExportObjectRef', 'ImportObjectRef'],
    incUsers.map(u => ['True', tf_(u['Priority?']), 'User', 'User', lc(u['[Source] Original Google Account']),
      lc(u['[Destination] Temp Google Account']), u['First Name'], u['Last Name'], pw[lc(u['[Destination] Temp Google Account'])],
      '', '', tf_(u['Email']), tf_(u['Contacts']), tf_(u['Calendar']), tf_(u['Drive']), 'False', 'False', '', '']));

  // 2. CloudM shared drives (ImportName = destination drive NAME; replace with the drive ID once the drives exist)
  files['cloudm_shared_drives.csv'] = toCsv_(
    ['Migrate', 'Priority', 'ExportObjectType', 'ImportObjectType', 'ExportName', 'ImportName', 'DocumentsPath', 'DocumentsDestinationPath'],
    incDrives.map(d => ['True', 'False', 'TeamDrive', 'TeamDrive', d['[Source] Shared Drive ID'], d['[Destination] Shared Drive Name'], '', '']));

  // 3. CloudM address replacements: no header, 2 columns, lower-case, unique source, no commas
  const seen = {}, ar = [];
  const add = (s, t, sec) => {
    s = lc(s); t = lc(t);
    if (!s || !t) return;
    if (s.indexOf(',') >= 0 || t.indexOf(',') >= 0) { warn.push(`Address replacement skipped (contains a comma): ${s}`); return; }
    if (seen[s] && seen[s] !== t) { warn.push(`Duplicate address replacement for ${s}: ${seen[s]} vs ${t}. The first one was kept.`); return; }
    if (!seen[s]) { seen[s] = t; ar.push([sec, s, t]); }
  };
  incUsers.forEach(u => add(u['[Source] Original Google Account'], u['[Destination] Temp Google Account'], 'Users'));
  incGroups.forEach(g => add(g['[Source] Group email'], g['[Destination] Temp group email'], 'Groups'));
  let arText = `# Generated ${new Date().toISOString()} from "${rv.getName()}"\n`;
  let sec = null;
  ar.forEach(([s0, s, t]) => { if (s0 !== sec) { arText += `# ${s0}\n`; sec = s0; } arText += `${s},${t}\n`; });
  if (String(cfg.domainCatchAll).toLowerCase().indexOf('y') === 0) {
    const dm = {};
    ar.forEach(([, s, t]) => { if (!dm[s.split('@')[1]]) dm[s.split('@')[1]] = t.split('@')[1]; });
    arText += '# Domain mappings (catch-all for internal addresses not listed above)\n';
    Object.keys(dm).sort().forEach(d => { if (d !== dm[d] && !seen[d]) arText += `${d},${dm[d]}\n`; });
  }
  files['cloudm_address_replacements.csv'] = arText;

  // 4. Org units (full source structure, parents before children)
  const ous = readRaw_('orgunits').filter(o => o.orgUnitPath !== '/')
    .sort((a, b) => a.orgUnitPath.split('/').length - b.orgUnitPath.split('/').length || a.orgUnitPath.localeCompare(b.orgUnitPath));
  files['dest_org_units.csv'] = toCsv_(['orgUnitPath', 'name', 'parentOrgUnitPath', 'description'],
    ous.map(o => [o.orgUnitPath, o.name, o.parentOrgUnitPath || '/', o.description]));

  // 5. Destination users
  files['dest_users.csv'] = toCsv_(
    ['email', 'firstName', 'lastName', 'password', 'orgUnitPath', 'sourceLicence', 'sourceEmail', 'finalEmail', 'sourceAfterCutover',
     'sourceGmailGB', 'sourceDriveGB', 'sourceGmailPlusDriveGB'],
    incUsers.map(u => [lc(u['[Destination] Temp Google Account']), u['First Name'], u['Last Name'],
      pw[lc(u['[Destination] Temp Google Account'])], u['OU Path'] || '/', u['Google Workspace licence'],
      lc(u['[Source] Original Google Account']), lc(u['[Destination] Final Google Account']), lc(u['[Source] Email address after cutover']),
      num_(u['Gmail (GB)']), num_(u['Drive (GB)']), num_(u['Gmail + Drive (GB)'])]));

  // 6. Destination groups (+ settings, with the client's external-posting decision applied)
  files['dest_groups.csv'] = toCsv_(
    ['email', 'name', 'description', 'sourceEmail'].concat(GROUP_SETTING_COLS),
    incGroups.map(g => {
      const s = GROUP_SETTING_COLS.map(k => g[k]);
      const ext = String(g['Do you need external users to send emails to the group?']).trim();
      const pi = GROUP_SETTING_COLS.indexOf('whoCanPostMessage');
      if (ext === 'Yes') s[pi] = 'ANYONE_CAN_POST';
      if (ext === 'No' && s[pi] === 'ANYONE_CAN_POST') s[pi] = 'ALL_IN_DOMAIN_CAN_POST';
      return [lc(g['[Destination] Temp group email']), g['Group Name'], g['Description'], lc(g['[Source] Group email'])].concat(s);
    }));

  // 7. Group members (only migrated groups; members excluded from migration are skipped)
  const incGroupSrc = new Set(incGroups.map(g => lc(g['[Source] Group email'])));
  const gmRows = [];
  gmembers.forEach(m => {
    if (!incGroupSrc.has(lc(m['Group email address']))) return;
    const dst = lc(m['[Destination] Member email']), g = m['Group email address'];
    if (lc(m['Group member email address']) === 'all_users_in_org') { warn.push(`${g}: 'all users in organization' member must be re-added manually after cutover.`); return; }
    if (isPlaceholder(dst) && !incTemp.has(dst)) {
      warn.push(`${g}: member ${m['Group member email address']} skipped (${allTemp.has(dst) ? 'excluded from migration' : 'not a current user or group'}).`);
      return;
    }
    gmRows.push([lc(m['[Destination] Temp group email']), dst, String(m['Member role']).toUpperCase(),
                 String(m['Member type']).indexOf('Group') === 0 ? 'GROUP' : 'USER']);
  });
  files['dest_group_members.csv'] = toCsv_(['groupEmail', 'memberEmail', 'role', 'type'], gmRows);

  // 7b. Aliases, one per row. Alias/Target/TargetType match GAM's own alias terms:
  //     gam csv dest_aliases.csv gam create alias "~Alias" "~TargetType" "~Target"
  const aliasRows = [], aliasSeen = {};
  const primaries = new Set(Array.from(allTemp));
  aliases.forEach(a => {
    if (String(a['Alias type']).indexOf('User-created') !== 0) return;          // automatic aliases are recreated by Google
    const src = a['[Source] Alias'], tgt = lc(a['[Destination] Target']), dst = lc(a['[Destination] Alias']);
    if (String(a['Migrate alias?']).trim() !== 'Yes') return;
    if (!incTemp.has(tgt)) { warn.push(`Alias ${src} skipped: its ${String(a['Target type']).toLowerCase()} ${a['[Source] Target (user or group)']} is not being migrated.`); return; }
    if (!dst || dst.indexOf('@') < 0) { warn.push(`Alias ${src} skipped: no destination alias address.`); return; }
    if (primaries.has(dst)) { warn.push(`Alias ${dst} skipped: it is also the address of a user or group.`); return; }
    if (aliasSeen[dst]) { warn.push(`Alias ${dst} skipped: listed more than once (targets ${aliasSeen[dst]} and ${tgt}).`); return; }
    aliasSeen[dst] = tgt;
    aliasRows.push([dst, tgt, String(a['Target type']).toLowerCase() === 'group' ? 'group' : 'user', lc(src), lc(a['[Source] Target (user or group)'])]);
  });
  files['dest_aliases.csv'] = toCsv_(['Alias', 'Target', 'TargetType', 'SourceAlias', 'SourceTarget'], aliasRows);

  // 8. Destination shared drives + ACLs
  files['dest_shared_drives.csv'] = toCsv_(['name', 'sourceDriveId', 'sourceName'],
    incDrives.map(d => [d['[Destination] Shared Drive Name'], d['[Source] Shared Drive ID'], d['[Source] Shared Drive Name']]));
  const incDriveIds = {};
  incDrives.forEach(d => { incDriveIds[d['[Source] Shared Drive ID']] = d['[Destination] Shared Drive Name']; });
  const aclRows = [];
  dacl.forEach(a => {
    const id = a['[Source] Shared Drive ID'];
    if (!(id in incDriveIds)) return;
    const dst = lc(a['[Destination] Collaborator email address']), ctype = String(a['Collaborator type']);
    const role = ROLE_FROM_LABEL[String(a['Role']).trim()];
    if (ctype === 'Domain' || ctype === 'Anyone') { warn.push(`${incDriveIds[id]}: '${a['[Source] Collaborator']}' permission not exported. Decide the destination target manually.`); return; }
    if (!role) { warn.push(`${incDriveIds[id]}: ${dst} has an unknown role "${a['Role']}". Skipped.`); return; }
    if (isPlaceholder(dst) && !incTemp.has(dst)) {
      warn.push(`${incDriveIds[id]}: ${a['[Source] Collaborator']} skipped (${allTemp.has(dst) ? 'excluded from migration' : 'not a current user or group'}).`);
      return;
    }
    aclRows.push([incDriveIds[id], id, dst, ctype.indexOf('Group') === 0 ? 'group' : 'user', role, a['Role']]);
  });
  files['dest_shared_drive_acl.csv'] = toCsv_(['driveName', 'sourceDriveId', 'emailAddress', 'type', 'role', 'roleLabel'], aclRows);
  incDrives.forEach(d => {
    if (!aclRows.some(r => r[1] === d['[Source] Shared Drive ID'] && r[4] === 'organizer'))
      warn.push(`${d['[Destination] Shared Drive Name']}: no Manager in the exported ACL. Assign one before creating the drive.`);
  });

  // 9. README
  const hasUsage = incUsers.some(u => num_(u['Gmail + Drive (GB)']) !== '');
  const sumGB = k => Math.round(incUsers.reduce((t, u) => t + (Number(num_(u[k])) || 0), 0) * 100) / 100;
  const storageLine = hasUsage
    ? `Source storage of included users (Gmail + Drive): ${sumGB('Gmail + Drive (GB)')} GB (Gmail ${sumGB('Gmail (GB)')} GB, Drive ${sumGB('Drive (GB)')} GB). ` +
      'Shared drive sizes are not available from Google APIs; use the CloudM scan.'
    : 'Source storage: not available in this review.';
  files['README.txt'] = exportReadme_(cfg, rv, signer, incUsers.length, incGroups.length, incDrives.length, warn, storageLine);

  // ---- zip -> Drive
  const stamp = Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyyMMdd-HHmm');
  const safe = cfg.project.replace(/[^\w.-]+/g, '_');
  const blobs = Object.keys(files).map(n => Utilities.newBlob(files[n], 'text/csv', n));
  const zip = Utilities.zip(blobs, `${safe}_migration_files_${stamp}.zip`);
  const file = Drive.Files.create({ name: zip.getName(), mimeType: 'application/zip' }, zip, { fields: 'id,webViewLink' });
  const props = PropertiesService.getScriptProperties();
  const ids = JSON.parse(props.getProperty(EXPORT_IDS_KEY) || '[]'); ids.push(file.id);
  props.setProperty(EXPORT_IDS_KEY, JSON.stringify(ids));
  const dl = `https://drive.google.com/uc?export=download&id=${file.id}`;
  setSetting_('exportUrl', dl);
  setSetting_('status', `Exported ${stamp}: ${incUsers.length} users, ${incGroups.length} groups, ${incDrives.length} shared drives, ${warn.length} warnings.`);

  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial;font-size:14px;line-height:1.5">
       <p><b>Export ready.</b> ${incUsers.length} users · ${incGroups.length} groups · ${incDrives.length} shared drives</p>
       <p><a href="${dl}" target="_blank" style="font-size:16px">⬇ Download ${zip.getName()}</a></p>
       <p>${warn.length ? `<b>${warn.length} warning(s)</b> are listed in README.txt inside the zip.` : 'No warnings.'}</p>
       <p style="color:#B00020">The zip contains temporary passwords. Send it to engineering through a secure channel, then run
       <i>Migration Mapping → Delete exported files</i>.</p></div>`).setWidth(480).setHeight(280);
  ui.showModalDialog(html, 'Migration Mapping');
}

function menuDeleteExports() {
  const props = PropertiesService.getScriptProperties();
  const ids = JSON.parse(props.getProperty(EXPORT_IDS_KEY) || '[]');
  let n = 0;
  ids.forEach(id => { try { Drive.Files.update({ trashed: true }, id); n++; } catch (e) { /* already gone */ } });
  props.deleteProperty(EXPORT_IDS_KEY);
  setSetting_('exportUrl', '');
  SpreadsheetApp.getUi().alert(`${n} exported file(s) moved to Trash.`);
}

// ------------------------------------------------------------------------------------------ helpers

function readTab_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return null;
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  return v.slice(1).filter(r => r.some(c => c !== '' && c !== null))
    .map(r => { const o = {}; h.forEach((k, i) => { o[k] = typeof r[i] === 'string' ? r[i].trim() : r[i]; }); return o; });
}

function signOffValue_(ss, label) {
  const sh = ss.getSheetByName(T.readme);
  if (!sh) return '';
  const v = sh.getDataRange().getValues();
  const row = v.find(r => String(r[0]).trim() === label);
  return row ? String(row[1] || '').trim() : '';
}

/** Number from a review cell ('' if blank or not numeric). */
function num_(v) { return v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v); }

function isYesNo_(v) { return v === 'Yes' || v === 'No'; }
function tf_(v) { return v === true || truthy_(v) ? 'True' : 'False'; }

function toCsv_(headers, rows) {
  const esc = v => { const s = v === null || v === undefined ? '' : String(v); return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  return [headers].concat(rows).map(r => r.map(esc).join(',')).join('\r\n') + '\r\n';
}

function genPassword_(n) {
  const alpha = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789', sym = '!@#$%^*-_';
  const pick = set => set.charAt(randInt_(set.length));
  while (true) {
    let p = '';
    for (let i = 0; i < n - 2; i++) p += pick(alpha);
    p += pick(sym) + pick('23456789');
    if (/[a-z]/.test(p) && /[A-Z]/.test(p)) return p;
  }
}

/** Unbiased random integer in [0, max) from the random bits of a v4 UUID (avoids Math.random for passwords).
 *  Google doesn't document Utilities.getUuid()'s entropy source, so treat these as temporary passwords only
 *  and force a change at first sign-in. */
function randInt_(max) {
  while (true) {
    const hex = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    const x = parseInt(hex, 16);              // 0 .. 2^32-1
    const lim = Math.floor(4294967296 / max) * max;
    if (x < lim) return x % max;
  }
}

function exportReadme_(cfg, rv, signer, nu, ng, nd, warn, storageLine) {
  return [
    `${cfg.project}: migration files`,
    `Generated: ${new Date().toISOString()}`,
    `From review: ${rv.getName()}  (${rv.getUrl()})`,
    `Approved by: ${signer || '(sign-off not completed)'}`,
    `Included: ${nu} users, ${ng} groups, ${nd} shared drives`,
    storageLine,
    '',
    'CLOUDM MIGRATE',
    '  cloudm_users.csv                 Bulk Add/Import Items: users (source -> placeholder). CONTAINS TEMP PASSWORDS.',
    '  cloudm_shared_drives.csv         Bulk Add/Import Items: shared drives (TeamDrive). ImportName is the destination',
    '                                   drive NAME. Replace it with the drive ID once the drives exist (a name makes CloudM create the drive).',
    '  cloudm_address_replacements.csv  Address Replacements (no header). Maps source users/groups to placeholders.',
    '',
    'DESTINATION BUILD (1:1 copy of the source on placeholder addresses)',
    '  dest_org_units.csv               OUs to create, parents first.',
    '  dest_users.csv                   Users to create (same passwords as the CloudM file), OU, source licence, source storage (GB).',
    '  dest_groups.csv                  Groups to create, with source settings.',
    '  dest_aliases.csv                 User and group aliases, one per row (Alias, Target, TargetType). For GAM:',
    '                                   gam csv dest_aliases.csv gam create alias "~Alias" "~TargetType" "~Target"',
    '  dest_group_members.csv           Group memberships and roles (OWNER/MANAGER/MEMBER).',
    '  dest_shared_drives.csv           Shared drives to create.',
    '  dest_shared_drive_acl.csv        Shared drive members (Drive API roles: organizer, fileOrganizer, writer, commenter, reader).',
    '',
    `WARNINGS (${warn.length})`,
  ].concat(warn.length ? warn.map(w => '  - ' + w) : ['  none']).join('\r\n') + '\r\n';
}
