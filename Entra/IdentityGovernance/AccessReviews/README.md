# Access reviews

[Create-AccessReview.ps1](./Create-AccessReview.ps1) is the unified entry point
for direct group-membership reviews. Keep
[AccessReview.Logging.ps1](./AccessReview.Logging.ps1) beside it.

## Input modes

Choose exactly one:

- `-GroupId`: one group, optionally with `-ReviewName`.
- `-CsvPath`: one review per CSV row; duplicate group IDs are processed once.
- `-ServicePrincipalId`: discover groups assigned to an enterprise application.
  Multiple application-role assignments for a group are collapsed. Direct user
  assignments are not reviewed. Names are `AR <application> - <group>`.

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

| ReviewerType | Below 500 transitive members | 500+ |
| --- | --- | --- |
| `ApplicationOwner` | Application owners | Application owners |
| `LineManager` | Managers -> Application owners | Self-review -> Managers -> Application owners |
| `ResourceOwner` | Group user owners -> Application owners | Self-review -> Group user owners -> Application owners |

If the group-owner (or explicit fallback) reviewer set exactly matches the
application-owner set, those stages are combined into one application-owner
stage. Order, duplicate IDs and letter casing do not affect the comparison;
partial overlap does not merge stages. Self-review remains separate at 500+
members. Dynamic per-user manager routing is not merged, even if its fallback
matches the application owners. Duration, stage dependencies and logs use the
resulting stages, and a warning reports the merge.

The 50/150-person cutoffs are removed in the unified behavior. Group counts
include transitive directory members, not only users. Configure
`$SelfReviewStageThreshold` (default 500, inclusive) and `$StageDurationInDays`
(default five) in the script. Group user owners are resolved at creation time;
service-principal owners are excluded.

If a group has no user owners, the ResourceOwner stage is omitted by default:
application owners review directly (500+ members still start with self-review).
To keep a separate first reviewer, supply `-FallbackReviewerUserId` with a
nonempty **user object ID**:

```powershell
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner `
    -FallbackReviewerUserId '44444444-4444-4444-4444-444444444444' -WhatIf
```

This produces fallback reviewer -> application owners for ownerless groups
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
For one-time reviews, EndDate is ignored; duration is stage count times days
per stage. Supply a later EndDate when needed for a recurring series.

Notifications and reminders are enabled; reviewer justification is required.
Unanswered reviews default to Deny, and denied access is automatically removed.
`$ReviewerMessage` adds optional email text; the full reviewer message supports
256 characters. `$AdditionalNotificationRecipients` optionally adds completion
notification recipients (user/group IDs), or leave it empty.

Synchronized groups are report-only: decisions are recorded but removal must
be performed in on-premises Active Directory. Automatic apply is disabled.

Application-discovery mode preserves its skip rules: dynamic groups, groups
with neither members nor user owners, and existing review names are skipped.
Existing reviews are not updated. Direct-group/CSV modes do not skip existing
names; check results before retrying to avoid duplicate reviews.

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

Every validated run appends JSON Lines to `logs/Create-AccessReview.jsonl`.
The whole `logs/` directory is gitignored. Entries contain UTC timestamps,
run ID, runnable invocation syntax (including array parameters), working
directory, effective schedule and per-group results including created IDs.
CSV files themselves are not copied into logs; effective owner IDs are logged.
Authentication commands, credentials and tokens are not logged. Keep logs
private because they contain tenant/group metadata.

WhatIf writes local preview logs. Logging errors stop execution explicitly;
check IDs already returned before retrying. Single-group failures are logged
before rethrowing; batch failures are reported per group. Application-discovery
failures are logged as `RunFailed`. Preflight errors occur before logging starts.

## Offline validation

```powershell
Invoke-Pester -Path '.\tests' -Output Detailed
```

Tests mock Graph; they never connect to or change a tenant.
