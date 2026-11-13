# Access reviews

[Create-AccessReview.ps1](./Create-AccessReview.ps1) is the unified entry point
for direct group-membership and direct application-user access reviews. Keep
[AccessReview.Logging.ps1](./AccessReview.Logging.ps1) beside it.

## Input modes

Choose exactly one:

- `-GroupId`: one group, optionally with `-ReviewName`.
- `-CsvPath`: group and/or enterprise-application object IDs in CSV rows.
  Application rows use the same discovery as `-ServicePrincipalId`.
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
    -ReviewerType ApplicationOwner -WhatIf
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

Direct-user reviews follow the same stage policy, schedule and notifications.
For `ResourceOwner`, actual service-principal **user** owners are the owner
stage for fewer than 500 direct users; non-user owners are never reviewers.
Without usable user owners, the target fails. At 500 or more direct users,
the final stage always uses explicit `AppOwnerIds`.

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

`-AppOwnerIds` accepts one or more **user object IDs** as run-wide defaults.
They are required whenever application owners review, including every target
with 500 or more members. They are optional when a smaller target uses
`ResourceOwner`; application-owner IDs are never inferred from discovery.

CSV rows contain either `GroupId` or `ServicePrincipalId`. Optional columns
are `AppOwnerIds`, `Owner`, `ReviewerType`, and `Role`; `ReviewName` is supported for
group rows only, because an application can produce multiple reviews:

```csv
GroupId,ServicePrincipalId,AppOwnerIds,ReviewerType,Role,ReviewName
11111111-1111-1111-1111-111111111111,,22222222-2222-2222-2222-222222222222,ResourceOwner,Finance approver,Finance review
,33333333-3333-3333-3333-333333333333,22222222-2222-2222-2222-222222222222,ApplicationOwner,Reader,
```

For a synchronized group, use `Owner` to supply one or more resource-owner
**user object IDs** without looking up group owners:

```csv
GroupId,ReviewerType,Owner
11111111-1111-1111-1111-111111111111,ResourceOwner,"22222222-2222-2222-2222-222222222222,33333333-3333-3333-3333-333333333333"
```

For a service-principal row, `Owner` applies to every synchronized group
discovered from that application, not to its direct-user review. The column is
used only for synchronized groups with `ResourceOwner` routing below 500
members; cloud groups use their Entra owners and 500+ reviews use application
owners. A populated `Owner` takes precedence over `-FallbackReviewerUserId`
for the owner stage. If `Owner` is blank, the explicit fallback users are used;
if neither is available, the target fails.

For an application row use the enterprise application's **object ID**, not
its client ID. Its role text and row settings apply to each group and direct-user
review discovered from that application. `Role` is added to the friendly
reviewer message; it does **not** filter the assignments or restrict access
being reviewed. Blank role cells add nothing.

Comma, semicolon (common in Excel exports), and tab column delimiters are
detected automatically from the header. Inside `AppOwnerIds` or `Owner`, separate multiple
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

Column order does not matter. Nonblank row owners and reviewer types replace
parameter defaults; blank/missing cells use the run-wide values. ReviewerType
must be `ApplicationOwner` or `ResourceOwner`. Duplicate identical targets are
processed once; conflicting duplicate row settings stop the run. Whitespace
around IDs is ignored. Malformed populated `Owner` values and missing or malformed required owners fail the affected
target without creating its review; other targets continue.
Owner IDs must identify users, not groups. Duplicate owner IDs are collapsed.
The script does not prevalidate that each GUID identifies an existing user.

## Stage selection

`-ReviewerType` (`-ReviewType` alias) selects the owner stage and defaults to
`ApplicationOwner`:

Empty groups are skipped with `No members; no access review needed`, even when
they have owners. This applies to single-group, CSV, and application discovery.
Applications without direct users do not get a direct-user review.

| Size | Stages |
| --- | --- |
| 1-150 | Selected owner |
| 151-499 | Line managers -> selected owner |
| 500+ | Self-review -> line managers -> application owners |

`ResourceOwner` selects group user owners or, for direct-user application reviews,
service-principal user owners. If no usable user owners exist on a non-synchronized
resource, the target fails even if a fallback is configured. For synchronized
groups below 500 members, CSV `Owner` users review instead of group owners.
If `Owner` is blank or absent, explicitly supplied fallback users review
instead; without either the target fails. At 500+ the owner selection is
ignored and application owners are required.
Group counts include transitive directory members, not only users; direct-user
application reviews count distinct directly assigned users.

`-FallbackReviewerUserId` accepts one or more user object IDs, for example:

```powershell
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner `
    -FallbackReviewerUserId '11111111-1111-1111-1111-111111111111','22222222-2222-2222-2222-222222222222' -WhatIf
```

Missing managers use these fallback users when supplied, otherwise the effective
application owners, or the selected resource owners when no application owners
were supplied. Duplicate IDs are removed. Fallback IDs are GUID-validated, not
checked for user existence. Lookup failures fail explicitly.
`-ReviewerType LineManager` is deprecated and maps to `ApplicationOwner` with
a warning; manager stages now depend on size. `-ForceTwoStage` is rejected,
because it conflicts with the required stage policy.
Graph advances `Approve` and `NotReviewed` decisions to subsequent stages;
denied decisions do not advance.

## Schedule and notifications

Defaults to `-Frequency OneTime` and `-StartDate` today. Recurring options:
`Weekly`, `Monthly`, `Quarterly`, `SemiAnnual`, `Annual`. Recurring `-EndDate`
defaults to the end of the current month and must allow all stages to finish.
For one-time reviews, EndDate is ignored; duration is the sum of the included
stages. Supply a later EndDate when needed for a recurring series.

Each stage defaults to five days. Set independent positive whole-day durations:

- `-SelfReviewDurationInDays`: self-review.
- `-SelectedReviewerDurationInDays`: managers.
- `-ApplicationOwnerDurationInDays`: the final owner stage, whether application
  owners or resource owners.

```powershell
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -AppOwnerIds '<application-owner user Object ID>' `
    -ReviewerType ResourceOwner `
    -SelfReviewDurationInDays 3 -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 5 -WhatIf
```

For 22 members this lasts 5 days; for 151-499 it lasts 12 days (7 + 5);
for 500+ it lasts 15 days (3 + 7 + 5). The same
parameters work in CSV and application modes. Recurring EndDate validation uses
the actual total, and real-run logs record all three configured durations.

Notifications and reminders are enabled; reviewer justification is required.
Unanswered reviews default to Deny, and denied access is automatically removed.
`$ReviewerMessage` adds optional email text; the full reviewer message supports
256 characters. `$AdditionalNotificationRecipients` optionally adds completion
notification recipients (user/group IDs), or leave it empty.

Synchronized groups are report-only: decisions are recorded but removal must
be performed in on-premises Active Directory. Automatic apply is disabled.
For `ResourceOwner` routing below 500 members, synchronized groups require CSV
owners or explicit fallback users as described above.

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
CSV files themselves are not copied into logs; effective application-owner IDs,
CSV owner IDs and role text are logged in results.
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
