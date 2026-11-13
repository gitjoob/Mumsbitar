# Access reviews

[Create-AccessReview.ps1](./Create-AccessReview.ps1) is the unified entry point
for direct group-membership and direct application-user access reviews. Keep
[AccessReview.Logging.ps1](./AccessReview.Logging.ps1) beside it.

## Input modes

Choose exactly one:

- `-GroupId`: one group, optionally with `-ReviewName`.
- `-CsvPath`: one review per CSV row; duplicate group IDs are processed once.
- `-ServicePrincipalId`: discover assigned groups, direct user assignments,
  service-principal owners, and other application-role assignments. Group reviews
  retain names `AR <application> - <group>`. Direct users receive an
  application-scoped review named `AR <application>`, without an assignment label
  or GUID. With only direct users, this is the only native review created. Owners and non-user
  assignments are explicitly reported for manual review (see below).

Use `-ReviewNamePrefix 'Custom'` in any mode to change the generated-name prefix
(default `AR`), or `-ReviewNamePrefix ''` for no prefix. Explicit `-ReviewName`
and CSV `ReviewName` values are used unchanged.

Connect to Graph first. Requires `Microsoft.Graph.Authentication`,
`Microsoft.Graph.Groups` and `Microsoft.Graph.Identity.Governance`;
application discovery also requires `Microsoft.Graph.Applications`.
Delegated sessions require `AccessReview.ReadWrite.All`, group read access
(including owners and transitive counts), and `Application.Read.All` for
application discovery. App-only authorization is enforced by Graph. Appropriate
Entra licensing and active administrator/resource roles are required as well.

```powershell
Connect-MgGraph -Scopes 'AccessReview.ReadWrite.All','Group.Read.All','Application.Read.All'

.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' `
    -AppOwnerIds '22222222-2222-2222-2222-222222222222' `
    -ReviewerType ResourceOwner -WhatIf

.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner -WhatIf

.\Create-AccessReview.ps1 -ServicePrincipalId '33333333-3333-3333-3333-333333333333' `
    -AppOwnerIds '22222222-2222-2222-2222-222222222222' `
    -ReviewerType LineManager -WhatIf
```

Replace example IDs with real object IDs. Remove `-WhatIf` after checking the
preview; preview performs Graph reads but never creates reviews.

## Application identity coverage

Application mode reads all pages of app-role assignments and service-principal
owners before creating any reviews. Discovery errors terminate the run rather
than silently producing incomplete coverage.

The console prints the resolved application name, service-principal Object ID,
distinct assigned group/user counts, and the direct-user review outcome (including
skip/failure reason). Discovery counts include users even when an existing review
causes creation to be skipped; the final created-review total does not.
Use the **enterprise application's Object ID**, not its application/client ID or
the app registration's Object ID. If Graph returns no direct users, check the
enterprise application's **Users and groups** assignments and allow for replication
after recent changes. To inspect the same read-only assignment source:

```powershell
Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId '<enterprise-application Object ID>' -All `
    -Property id,principalId,principalType,principalDisplayName,appRoleId,resourceId |
    Select-Object PrincipalType,PrincipalId,PrincipalDisplayName,Id,AppRoleId
```

| Relationship | Handling |
| --- | --- |
| Assigned groups | Existing group-membership reviews, with the existing skip rules. These do not review or remove the group's assignment to the application. |
| Directly assigned users | One native application-access review. Multiple roles for the same user count once in `SubjectCount`; discovered principal and assignment IDs are retained in results/logs. |
| Service-principal owners | Warning, structured result and real-run log entry marked `Manual review required`. Ownership is distinct from application access. |
| Assigned service principals (including managed identities) | Each role assignment is explicitly reported as `Manual review required`, retaining principal, assignment and app-role IDs. |
| Other/unknown assignment types | Explicit manual-review results, never silently discarded. |

The direct-user review uses the documented
[`principalResourceMembershipsScope`](https://learn.microsoft.com/en-us/graph/accessreviews-scope-concept#example-15-review-all-users-with-direct-or-transitive-assignment-to-an-application)
with `/v1.0/users` as its sole principal scope and the selected service principal
as its resource. The group-principal scope from that example is deliberately
omitted: group membership remains covered by the separate group reviews.
Its subject count is a discovery-time snapshot; Graph resolves the review scope
when instances start, including subsequent recurring instances.

Direct-user reviews retain the configured stage policy, schedule, notifications,
and final `AppOwnerIds` sign-off. For `ResourceOwner`, actual service-principal
**user** owners are first-stage reviewers; non-user owners are never reviewers.
Identical reviewer sets still merge. Without user owners, the same fallback/skip
policy applies. The 500-member self-review threshold uses distinct direct users.

**Denied direct-user reviews automatically remove application access**, not group
membership or ownership. This is separate from the existing group-membership
removal behavior. Start with `-WhatIf` to inspect both kinds of targets.

Manual-review results are an inventory requiring administrator follow-up, **not
native access-review definitions or completed reviews**. No ownership or workload
assignment is removed by the script. It does not invent unsupported owner or
workload scopes for the user-access review. The inventory covers the enterprise
application/service principal, not the separate app registration's owners,
delegated consent grants, directory roles, or every possible access path.
`AppOwnerIds` remains an explicit sign-off list, not a discovery filter.

## Application-owner sign-off and CSV

Application-owner sign-off is required in every mode. `-AppOwnerIds` accepts
one or more **user object IDs** as run-wide defaults. Application discovery
requires these explicit IDs; it never infers application owners.

CSV requires `GroupId`; `AppOwnerIds` and `ReviewName` are optional columns:

```csv
GroupId,AppOwnerIds,ReviewName
11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222;33333333-3333-3333-3333-333333333333,Finance review
44444444-4444-4444-4444-444444444444,55555555-5555-5555-5555-555555555555,Service desk review
```

Comma, semicolon (common in Excel exports), and tab column delimiters are
detected automatically from the header. Inside `AppOwnerIds`, separate multiple
IDs with commas or semicolons. Quote the cell when its separator matches the
CSV column delimiter. For example, comma-separated IDs in a comma-delimited file:

```csv
GroupId,AppOwnerIds
11111111-1111-1111-1111-111111111111,"22222222-2222-2222-2222-222222222222,33333333-3333-3333-3333-333333333333"
```

Comma-separated IDs also work in a semicolon-delimited file:

```csv
GroupId;AppOwnerIds
11111111-1111-1111-1111-111111111111;22222222-2222-2222-2222-222222222222,33333333-3333-3333-3333-333333333333
```

Column order does not matter. Nonblank row owners replace parameter defaults;
blank/missing cells use `-AppOwnerIds`. Whitespace around IDs is ignored. Missing or
malformed owners fail that row without creating its review; other rows continue.
Owner IDs must identify users, not groups. Duplicate owner IDs are collapsed.
The script does not prevalidate that each GUID identifies an existing user.

## Stage selection

`-ReviewerType` (`-ReviewType` alias) defaults to `LineManager`:

Empty groups are skipped with `No members; no access review needed`, even when
they have owners. This applies to single-group, CSV, and application discovery.
Applications without direct users do not get a direct-user review.

| ReviewerType | 1-149 members | 150-499 members | 500+ |
| --- | --- | --- | --- |
| `ApplicationOwner` | Application owners | Application owners | Application owners |
| `LineManager` | Application owners | Managers -> Application owners | Self-review -> Managers -> Application owners |
| `ResourceOwner` | Application owners | Group user owners -> Application owners | Self-review -> Group user owners -> Application owners |

For special cases, use `-ForceTwoStage` with `-ReviewerType ResourceOwner` or
`LineManager` to include the selected reviewer stage even below 150 members:

```powershell
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -AppOwnerIds '<application-owner user Object ID>' `
    -ReviewerType ResourceOwner -ForceTwoStage -WhatIf
```

For a 22-member group, this creates **group owners -> application owners** instead
of application owners only. Without the switch, the defaults above are unchanged.
The switch applies to single-group, CSV, and application modes. It forces at least
two stages, retaining self-review at 500+ members (three stages). Identical owner
sets are not merged when forced. Ownerless resources require
`-FallbackReviewerUserId`; otherwise the target fails explicitly rather than
silently creating one stage. Empty groups are still skipped.
`-ReviewerType ApplicationOwner -ForceTwoStage` is rejected as contradictory.
Use `-WhatIf` to preview; real-run logs record the override.

If the group-owner (or explicit fallback) reviewer set exactly matches the
application-owner set, those stages are combined into one application-owner
stage unless `-ForceTwoStage` is supplied. Order, duplicate IDs and letter casing do not affect the comparison;
partial overlap does not merge stages. Self-review remains separate at 500+
members. Dynamic per-user manager routing is not merged, even if its fallback
matches the application owners. Duration, stage dependencies and logs use the
resulting stages, and a warning reports the merge.

The 150-member threshold applies to all modes, including direct application
users. Below it, group-owner lookups and fallback stages are not needed unless
`-ForceTwoStage` is supplied. Group counts
include transitive directory members, not only users. Configure
`$SelfReviewStageThreshold` (default 500, inclusive) in the script.
Stage durations are configured with the parameters below. Group user owners are resolved at creation time;
service-principal owners are excluded.

If a group has no user owners, the ResourceOwner stage is omitted by default:
application owners review directly (500+ members still start with self-review).
To keep a separate first reviewer, supply `-FallbackReviewerUserId` with a
nonempty **user object ID**:

```powershell
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner `
    -FallbackReviewerUserId '44444444-4444-4444-4444-444444444444' -WhatIf
```

At 150+ members, this produces fallback reviewer -> application owners for ownerless groups
(preceded by self-review at 500+ members). Stage counts, duration and logs reflect
the actual stages. The script warns whenever it substitutes or omits group owners.
Missing managers use the same explicit fallback, or the effective application
owners for that CSV row when no fallback is supplied. No zero-GUID placeholder
is sent as a reviewer. Fallback IDs are GUID-validated, not checked for user existence.
Lookup failures fail explicitly rather than silently using fallback.
Graph advances `Approve` and `NotReviewed` decisions to subsequent stages;
denied decisions do not advance. Final-stage review is by application owners.

## Schedule and notifications

Defaults to `-Frequency OneTime` and `-StartDate` today. Recurring options:
`Weekly`, `Monthly`, `Quarterly`, `SemiAnnual`, `Annual`. Recurring `-EndDate`
defaults to the end of the current month and must allow all stages to finish.
For one-time reviews, EndDate is ignored; duration is the sum of the included
stages. Supply a later EndDate when needed for a recurring series.

Each stage defaults to five days. Set independent positive whole-day durations:

- `-SelfReviewDurationInDays`: self-review.
- `-SelectedReviewerDurationInDays`: managers, resource owners, or their explicit
  fallback stage.
- `-ApplicationOwnerDurationInDays`: final application-owner sign-off, including
  application-owner-only reviews.

```powershell
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -AppOwnerIds '<application-owner user Object ID>' `
    -ReviewerType ResourceOwner -ForceTwoStage `
    -SelfReviewDurationInDays 3 -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 5 -WhatIf
```

For 22 members with forced two-stage routing, this lasts 12 days (7 + 5).
For 500+ members it lasts 15 days (3 + 7 + 5). Skipped or merged stages add no
time; a merged owner stage uses only the application-owner duration. The same
parameters work in CSV and application modes. Recurring EndDate validation uses
the actual total, and real-run logs record all three configured durations.

Notifications and reminders are enabled; reviewer justification is required.
Unanswered reviews default to Deny, and denied access is automatically removed.
`$ReviewerMessage` adds optional email text; the full reviewer message supports
256 characters. `$AdditionalNotificationRecipients` optionally adds completion
notification recipients (user/group IDs), or leave it empty.

Synchronized groups are report-only: decisions are recorded but removal must
be performed in on-premises Active Directory. Automatic apply is disabled.
For `ResourceOwner` routing, synchronized groups use `-FallbackReviewerUserId`
instead of querying group owners. Without an explicit fallback, the effective
`AppOwnerIds` are used. Identical reviewer stages still merge unless
`-ForceTwoStage` is set. The below-150 default and empty-group skips are unchanged.

All modes skip an existing review with the same display name (case-insensitive),
including completed reviews; existing reviews are not updated. The script stops
without creating reviews if it cannot retrieve existing names. A preview also
skips duplicate names within the batch. Matching is by name, not resource ID, so
renaming a review or changing its prefix can allow another review for the same
resource. Concurrent runs can still create duplicates.
Application-discovery mode also skips dynamic groups; all modes skip empty groups.

## Results and private logs

Results include group/review IDs, `MemberCount`, effective application-owner IDs, stage names,
reviewer type, status and failure reason. Batch summaries are written separately
from result objects. A total transitive directory-member count appears directly
below the batch summary and is also printed for single-group runs. It includes
only created reviews, or would-be-created reviews with `-WhatIf`; skipped,
failed, and declined reviews are excluded. The total is zero if no reviews
qualify. Shared members count separately for each group. Counts include nested
membership and directory objects other than users; reviews themselves still
cover direct membership. `MemberCount` is null if the count was not retrieved,
and is also retained in per-group logs:

```powershell
$results = .\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner
$results | Export-Csv '.\results.csv' -NoTypeInformation
```

Application-mode results also identify `TargetType` and `ServicePrincipalId`.
`ApplicationDirectUsers` results have `SubjectCount`, `PrincipalIds` and
`AssignmentIds` instead of `MemberCount`. Ownership and non-user results have
`PrincipalId`, `PrincipalType`, `PrincipalDisplayName`, and where applicable
`AssignmentId` and `AppRoleId`. All result types expose the same fields for
mixed-result CSV exports. Application summaries count native review targets
separately from manual-review items and print direct-user totals separately from
group transitive-member totals.

Every validated real run writes a separate JSON Lines file named
`logs/Create-AccessReview_<UTC timestamp>_<RunId>.jsonl`, for example
`logs/Create-AccessReview_20261110T083000123Z_11111111-1111-1111-1111-111111111111.jsonl`.
The timestamp makes runs easy to sort; the unique RunId keeps runs isolated.
The script prints the log path when logging starts, including for runs that
later fail. Events are written as they occur, so partial logs remain available.
Real runs retain diagnostic logs even if no reviews are created.
The whole `logs/` directory is gitignored. Entries contain UTC timestamps,
run ID, runnable invocation syntax (including array parameters), working
directory, effective schedule and per-group results including created IDs.
Application reviews use `ApplicationResult` events; manual-review inventory uses
`ManualReviewRequired` events and a separate count in `RunCompleted`.
CSV files themselves are not copied into logs; effective owner IDs are logged.
Authentication commands, credentials and tokens are not logged. Keep logs
private because they contain tenant/group metadata.

WhatIf creates no log files or directories and does not change existing logs.
All real-run logs are kept until you remove them manually. The old combined
`logs/Create-AccessReview.jsonl` is left untouched; it is not migrated or appended to.

Open the printed path to inspect one run, or find the most recent run:

```powershell
$log = Get-ChildItem '.\logs\Create-AccessReview_*.jsonl' | Sort-Object Name -Descending | Select-Object -First 1
Get-Content -LiteralPath $log.FullName | ConvertFrom-Json
```

Logging errors stop real runs explicitly;
check IDs already returned before retrying. Single-group failures are logged
before rethrowing; batch failures are reported per group. Application-discovery
failures are logged as `RunFailed`. Preflight errors occur before logging starts.

## Offline validation

```powershell
Invoke-Pester -Path '.\tests' -Output Detailed
```

Tests mock Graph; they never connect to or change a tenant.
