# Google-to-Google Mapping Builder v2.2.0 (cloud-only)

Runs entirely in Google Sheets inside the **source** tenant. Nothing is installed locally and nothing is collected from the
destination. The destination is a 1:1 rebuild of the source on placeholder addresses (`user@domain` → `user@temp.domain`).

```
single-file/appsscript.json  manifest: scopes + advanced services (Admin SDK Directory, Groups Settings, Drive v3)
single-file/Code.gs          the three .gs files from v2 merged, for the paste install (+ appsscript.json)
```

## Installing the tool in a customer tenant (paste install, about 3 minutes)

1. **PM, own browser profile (Cloudasta account):** open the internal install kit (`single-file` folder), which contains the `Code.gs` and
   `appsscript.json` files.
2. **PM, on a separate browser profile signed in as the customer super admin:** create a blank Google Sheet → Extensions → Apps Script.
3. Click on Project Settings (gear icon over the left) → and tick **Show "appsscript.json" manifest file in editor**.
4. Editor: replace the contents of `appsscript.json` with those from the provided file. Then, add a new script file named `Code.gs` and replace the contents as well, once done, save the changes. Reload the Sheet and the **Migration Mapping** menu appears.

The script is created by the customer's own admin, so it is an internal app: no OAuth verification and no unverified-app cap.

PM steps are inside the tool (Instructions tab): Settings → 1. Collect → 2. Build review → client signs off → 3. Export → Delete exported files.

## What engineering receives (zip)
| File | Purpose |
|---|---|
| `cloudm_users.csv` | CloudM Bulk Add/Import Items: users. **Contains temp passwords** |
| `cloudm_shared_drives.csv` | CloudM items, `TeamDrive`. ImportName = destination drive name. Swap it for the drive ID once created |
| `cloudm_address_replacements.csv` | CloudM Address Replacements: no header, lower-case, unique sources, plus a domain catch-all |
| `dest_org_units.csv` | Full source OU tree, parents first |
| `dest_users.csv` | Placeholder users: same passwords as the CloudM file, OU, aliases mapped to the temp domain |
| `dest_groups.csv` | Placeholder groups and their source settings. The client's external-posting answer is applied |
| `dest_group_members.csv` | Members and roles (OWNER/MANAGER/MEMBER). Excluded and deleted members are skipped |
| `dest_shared_drives.csv`, `dest_shared_drive_acl.csv` | Drives to create and their members (Drive API roles) |
| `README.txt` | Approver, counts, and every warning raised during export |

The `dest_*` files are plain CSVs, so you can use them with GAM in the destination or with a provisioning script. If useful,
the next step is a "Provision destination" menu for the destination tenant. It would create OUs, users, groups, memberships, drives and ACLs from this zip.

## What's new in v2.2.0
- **Google Photos removed.** Photos is no longer requested from the Reports API.
  - The per-user total is now **Gmail + Drive (GB)**.
  - The tenant total is **Gmail + My Drive + shared drives**: what we migrate.
  - **Quota used % (all Google storage)** is still Google's own figure (it includes Photos), because it shows real quota pressure. The 80% flag uses it.
- **Aliases tab** (user and group aliases), one per row with the target first:
  - **Columns:** Target, Alias, Target type, Alias type, destination Target and Alias, Migrate alias? (client), Flags.
  - **Automatic aliases** (from domain aliases) are listed for information only. They're marked N/A and not exported.
  - **Elsewhere:** the alias column is gone from User Mapping, and Groups (Summary) shows an alias count instead of a list.
- **Licence column.** User Mapping now shows the **Google Workspace licence** where the alias column was.
  - **Source:** the Enterprise License Manager API (`AdminLicenseManager`, scope `apps.licensing`). I found no read-only scope for licensing; the tool only reads.
  - **Products queried:** Workspace (`Google-Apps`), Archived User (`101034`) and Cloud Identity (`101001`/`101005`), using the customer ID (`C0…`).
  - **Shown:** the Workspace plan, else the Archived User licence, else Cloud Identity, else "No licence".
  - **Flags:** active users with no licence. A licence count goes on the Read Me tab.
  - **Products the tenant doesn't have** are recorded and skipped. If licensing fails entirely, the column is blank and an Info issue is raised.
- **New Info issues:** more than one source domain (lists every `temp.` domain the destination needs), and automatic aliases present.
- **Export:**
  - **New `dest_aliases.csv`** (Alias, Target, TargetType, SourceAlias, SourceTarget), ready for
    `gam csv dest_aliases.csv gam create alias "~Alias" "~TargetType" "~Target"`.
    Aliases of excluded users and groups, duplicates, and aliases that clash with a primary address are skipped with a warning.
  - **`dest_users.csv`:** loses `aliases`, gains `sourceLicence`; `sourceTotalGB` becomes `sourceGmailPlusDriveGB`.
  - **`dest_groups.csv`:** loses `aliases`.
- **Compatibility:** the export needs a review built by v2.2.0, because it reads the Aliases tab. Re-run step 2 for older reviews.
  Re-authorization is needed on the next run (one new permission).

## What's new in v2.1.0: storage usage
- **Collection** adds two phases, `tenantusage` and `userusage`, using the Admin SDK Reports API (`AdminReports` advanced service,
  scope `admin.reports.usage.readonly`). This adds one item to the authorization screen. Collection walks back up to 7 days
  (Pacific time) to the newest date that has data, because Google's storage reports lag 1–3 days.
- **Review spreadsheet:**
  - The **Google User Mapping** tab gains **Gmail (GB), Drive (GB), Total used (GB) and Quota used %** (numeric, sortable).
  - A new **Storage Summary** tab shows tenant totals (total, Gmail, Drive, all shared drives, Photos, quota, users above 80% and
    50–80% of quota), storage by default Migrate? decision, and the 10 largest accounts.
  - The Read Me summary includes a storage line.
- **New flags:**
  - Users at **80% of quota or more**: flagged, with a Low issue.
  - Users missing from the report: flagged.
  - Reports API unavailable: an Info issue.
- **Export:** `dest_users.csv` gains `sourceGmailGB`, `sourceDriveGB` and `sourceTotalGB`. README.txt shows the total source storage
  of included users.
- **Fails safe:** if the Reports API is denied or returns nothing, collection still completes. The usage columns are left blank and
  the Status row says "storage usage unavailable".
- **Limits (Google API constraints):**
  - Sizes only. The Reports API has no per-user message or file counts.
  - No per-shared-drive size, only the all-drives total. Use the CloudM scan for per-drive sizes.
  - Figures are as of the report date, not live.

## Issues the tool can flag
- **Users:** suspended, archived, never signed in, inactive for more than N days, super admin, missing first or last name.
- **Groups:** no members, external members, nested groups, “all users in the organization” as a member, unreadable settings, custom access settings.
- **Shared drives:** no Manager, external collaborators, domain-wide or public sharing, deleted accounts in the access list, hidden drive, unreadable access list.

## Behaviour and limits
- **Resumable collection.** Apps Script caps each run at 6 minutes, so the tool saves its progress about every 4.5 minutes and
  resumes itself 1 minute later. Workspace accounts get 6 hours of trigger runtime per day, which is far more than collection needs.
- **The build and export steps are not resumable.** They run in memory and should finish well within 6 minutes for typical tenants.
  Very large tenants (tens of thousands of ACL rows) are untested.
- **Default `Migrate?`:** Yes for active users who have signed in. Review for suspended, archived and never-signed-in users.
  Export asks before treating unresolved rows as No, and warns if the sign-off is empty.
- **Group access type** is a best-effort label derived from Groups Settings ("Custom" when nothing matches). The settings
  columns are the source of truth.
- **Not collected:** email forwarding (needs domain-wide delegation) and file-level sharing inside drives (CloudM migrates
  that with the address replacements).
- **Drive-wide `domain`/`anyone` permissions** are flagged and left out of the ACL file. Someone has to decide the destination target.
- **Temp passwords** come from the random bits of `Utilities.getUuid()`. Google doesn't document that function's entropy source,
  so treat them as temporary only and force a change at first sign-in.
- **No collision checks against the destination** (by design). An existing placeholder address will show up as an error at
  provisioning time.

## Testing status
- ~~Initial testing of this version of the tool successfully retrieved information from the source tenant.~~
- ~~Most sections of the current project plan (PP) have been populated successfully.~~
- ~~A few details in the PM-to-ENG deliverables handoff still need to be polished.

## WIP
- ~~Add Email/Drive usage per user + tenant totals.~~
- ~~Add all source aliases in User Mapping sheet.~~
- ~~Add license per user (if possible).~~
- Details on the generated CSV files for engineering.
