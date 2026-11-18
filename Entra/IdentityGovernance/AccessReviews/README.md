# Access reviews

Create access reviews for groups and enterprise applications, or report on
existing reviews. Start with `-WhatIf` before creating anything.

## Quick start

Requires the Microsoft Graph PowerShell modules `Authentication`, `Groups` and
`Identity.Governance`; application discovery also needs `Applications`.
Use an account with the appropriate Entra role and access-review license.
Hidden memberships additionally require `Member.Read.Hidden`.

Keep [Create-AccessReview.ps1](./Create-AccessReview.ps1),
[AccessReview.Logging.ps1](./AccessReview.Logging.ps1) and
[AccessReview.Reporting.ps1](./AccessReview.Reporting.ps1) together.

```powershell
Connect-MgGraph -Scopes 'AccessReview.ReadWrite.All','Group.Read.All','Application.Read.All'

# One group and its nested groups; each uses its own group owners.
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -ReviewerType ResourceOwner -WhatIf

# Batch of groups and/or applications.
.\Create-AccessReview.ps1 -CsvPath '.\reviews.csv' -ReviewerType ResourceOwner -WhatIf

# Enterprise application: assigned groups plus a separate review of direct users.
.\Create-AccessReview.ps1 -ServicePrincipalId '<enterprise-application Object ID>' `
    -AppOwnerIds '<owner user Object ID>' -WhatIf
```

Replace the example IDs. For applications, use the **enterprise application's
Object ID**, not its client ID or app registration ID.
Remove `-WhatIf` only after checking the results.

## Reviewers and stages

Stages use each group's **direct-user count**, or the application's directly
assigned user count. Empty targets are skipped.

| Direct users | Stages |
| --- | --- |
| 1-150 | Owner |
| 151-499 | Line managers -> owner |
| 500+ | Self-review -> line managers -> owner |

- `-ReviewerType ApplicationOwner` is the default; supply `-AppOwnerIds`.
- `-ReviewerType ResourceOwner` uses the group's or application's human owners.
- `-ForceTwoStage` adds managers for 1-150 users; larger targets keep their stages.
- `-FallbackReviewerUserId` supplies fallback users for missing managers and
  ownerless groups. Without explicit manager fallbacks, supplied application
  owners are used, then selected resource owners.
- Synced groups use CSV `Owner` or explicit fallback users, not Entra group
  owners. A direct-user application review without human owners fails when
  `ResourceOwner` is selected.

## CSV

Each row must contain **either** `GroupId` **or** `ServicePrincipalId`.

| Column | Purpose |
| --- | --- |
| `GroupId` | Group Object ID. |
| `ServicePrincipalId` | Enterprise-application Object ID. |
| `ReviewerType` | Optional: `ApplicationOwner` or `ResourceOwner`. |
| `AppOwnerIds` | Optional application-owner user IDs; replaces parameter defaults. |
| `Owner` | Optional user IDs for synced groups using `ResourceOwner`. |
| `Role` | Optional text added to the reviewer message; does not filter access. |
| `ReviewName` | Optional custom name for group rows only. |

```csv
GroupId,ServicePrincipalId,ReviewerType,AppOwnerIds,Owner,Role,ReviewName
11111111-1111-1111-1111-111111111111,,ResourceOwner,,,Reader,Finance review
,33333333-3333-3333-3333-333333333333,ApplicationOwner,22222222-2222-2222-2222-222222222222,,Reader,
```

Comma, semicolon and tab delimiters are detected automatically. Separate multiple
owner IDs with commas or semicolons; quote cells when needed. Owner IDs must be
**user** Object IDs. Blank optional cells use the run-wide defaults where applicable.

## Nested groups and application coverage

- Each discovered group gets at most one review per run, covering its direct
  users. Shared groups retain all discovered paths.
- Children inherit reviewer type, application owners and role text, but **not**
  CSV `Owner` or `ReviewName`. Synced children need their own owner row or fallbacks.
- Empty or already-reviewed parents do not prevent discovering their children.
  Rerun discovery to include newly nested groups in future campaigns.
- Circular nesting produces a warning; that branch stops, other discovery continues.
  Membership-read failures or conflicting inherited settings stop creation.
- Application owners and non-user assignments are listed for **manual follow-up**,
  not automatically reviewed or removed. App registration ownership and consent
  grants are outside this tool's coverage.

Names default to `AR <group>` or `AR <application> - <assigned group>`.
Nested-only groups use `AR <group>`; direct-user application reviews use
`AR <application>`. Change the prefix with `-ReviewNamePrefix`.

## Schedule

Defaults: one-time review, starting today, **five days per stage**.

| Option | Meaning |
| --- | --- |
| `-StartDate` | First review start date. |
| `-EndDate` | For one-time reviews: explicit total review window, split across stages. For recurring reviews: series end date, defaulting to this month's end. |
| `-Frequency` | `OneTime`, `Weekly`, `Monthly`, `Quarterly`, `SemiAnnual` or `Annual`. |
| `-SelfReviewDurationInDays` | Self-review duration. |
| `-SelectedReviewerDurationInDays` | Manager duration. |
| `-ApplicationOwnerDurationInDays` | Final owner-stage duration, for either owner type. |

Allow at least one day per stage. For recurring reviews, supply an end date that
allows the intended cycles to run. Notifications and reminders are enabled.
Edit `$ReviewerMessage` for extra email text (256-character total limit), or
`$AdditionalNotificationRecipients` for completion notifications.

## Safety and duplicate protection

**Denied and unanswered access is removed automatically.**
Synced groups are report-only: apply removals in on-premises AD.
Removing someone from a shared group affects **every resource using that group**.
Group reviews remove user membership, not group-to-group or group-to-app assignments.

- Existing names and legacy aliases block creation, including completed reviews.
- Resource ID/scope checks also block active/scheduled or potentially future
  recurring coverage, regardless of prefix or reviewer policy.
- Completed historical coverage alone does not block by resource. Future/no-end
  recurring ranges are treated conservatively; numbered series check instances.
- Partial or unsupported overlapping scopes require manual assessment; no new
  review is created. Unreadable duplicate-check data stops creation.
- Existing reviews are **never changed, stopped or deleted**. Separate group
  reviews remain separate. Do not run overlapping creation jobs concurrently.

`-WhatIf` performs Graph reads but writes no reviews or logs. It writes an HTML
file only when you explicitly supply `-HtmlReportPath`.

## Reports and logs

### Creation report

```powershell
.\Create-AccessReview.ps1 -CsvPath '.\reviews.csv' -ReviewerType ResourceOwner `
    -WhatIf -HtmlReportPath '.\reports\preview.html'

# Convert a single run's JSONL log to HTML without connecting to Graph.
.\Export-AccessReviewReport.ps1 -LogPath '.\logs\<run>.jsonl' `
    -OutputPath '.\reports\creation.html'
```

The report includes a clickable group topology, stage configuration, duplicate
matches, search and a dark-mode switch. It shows planned/submitted settings and
creation outcomes, **not live status or proof that access was removed**.

Real runs write a separate JSONL file under `logs/` and print its path. Results
contain review/resource IDs, counts, discovery paths and skip/failure reasons.
Check returned IDs and logs before retrying a failed run. Logs are not deleted
automatically. To export the returned results:

```powershell
$results = .\Create-AccessReview.ps1 -CsvPath '.\reviews.csv' -ReviewerType ResourceOwner
$results | Export-Csv '.\results.csv' -NoTypeInformation
```

### Status report

[Get-AccessReviewStatusReport.ps1](./Get-AccessReviewStatusReport.ps1) reads
existing reviews and **all historical instances** matching literal name prefixes.
It requires the `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users`
modules, plus the shared reporting helper.

```powershell
Connect-MgGraph -Scopes 'AccessReview.Read.All','User.Read.All','GroupMember.Read.All','Application.Read.All'

.\Get-AccessReviewStatusReport.ps1 -Prefixes @('AR ', 'Quarterly ') `
    -HtmlReportPath '.\reports\status.html' -OpenReport
```

Includes stages, individual users, decisions, recorded reviewers and current
owners/managers. **Current directory relationships are not historical reviewer
assignments.** Manager lookups require `User.Read.All`. Missing data appears as
collection issues, not zero counts. The script makes no tenant changes.

All HTML reports are offline, have no external assets, and refuse to overwrite
files. Protect reports and logs: they contain private tenant/user information.

## Contributors

See [AGENTS.md](./AGENTS.md) for implementation rules. Tests mock Graph and never
change a tenant:

```powershell
Invoke-Pester -Path '.\tests' -Output Detailed
```
