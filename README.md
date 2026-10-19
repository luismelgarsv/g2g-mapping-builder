# Google-to-Google Mapping Builder v2.0.1 (cloud-only)

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
- Initial testing of this version of the tool successfully retrieved information from the source tenant.
- Most sections of the current project plan (PP) have been populated successfully.
- A few details in the PM-to-ENG deliverables handoff still need to be polished.
