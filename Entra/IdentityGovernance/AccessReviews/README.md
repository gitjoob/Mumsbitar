# Access reviews

[Create-AccessReview.ps1](./Create-AccessReview.ps1) is the unified entry point
for direct group-membership and direct application-user access reviews. Keep
[AccessReview.Logging.ps1](./AccessReview.Logging.ps1) beside it.

## Input modes

Choose exactly one:

- `-GroupId`: one root group, optionally with `-ReviewName`, plus its nested groups.
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
(including owners and direct membership), and `Application.Read.All` for
application discovery. App-only authorization is enforced by Graph. Appropriate
Entra licensing and active administrator/resource roles are required as well.
Hidden-membership groups also require `Member.Read.Hidden` and an appropriate
directory role. A discovery/permission error stops the run before any reviews
are created; it is never interpreted as an empty group.

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

## Nested groups and direct-user coverage

All modes automatically discover nested groups recursively before creating any
reviews. Each distinct group ID gets at most one group review per run, covering
its **direct users**. Stage selection uses that group's distinct direct-user
count, not its transitive membership or number of nested groups:

| Group | Direct membership | Review |
| --- | --- | --- |
| Group 1 | 50 users and Group 2 | 1 stage for the 50 direct users |
| Group 2 | 470 users | 2 stages for the 470 direct users |

`-ForceTwoStage` still adds managers for groups with 1-150 direct users.
With `-ReviewerType ResourceOwner`, each cloud group uses **its own human
owners**, including child groups with three-stage reviews.

Discovery continues through groups with zero direct users and groups whose
review names already exist. A rerun therefore skips existing parents but can
create missing child reviews. It does **not** correct the stages or duration of
an existing review. Nested groups are discovered again on each script run;
recurring review instances do not automatically discover new child groups.
Counts used for stage selection are a discovery-time snapshot.

### Names and shared children

Root names remain unchanged, including explicit `ReviewName` values and
`AR <application> - <group>` for application-assigned root groups.
Automatically discovered nested-only groups use `AR <child>` (or your
`ReviewNamePrefix`), independent of the order or number of parents.
The technical `[Nested]` suffix is no longer added. Explicit custom names are
preserved, and existing/running reviews are never renamed or updated.

If a child is also an explicit root in the current input, its root name takes
precedence. When several application roots reference the same group, an existing
root naming candidate is chosen deterministically. All discovery paths are
retained regardless of the name chosen.

The existing case-insensitive name guard also checks legacy
`<prefix> <child> [Nested]` names, including when that child is now a standalone
root, and applicable `<prefix> <application> - <child>` names,
plus explicit names supplied for that group in the current input. A match skips
creation, even for completed reviews. Unknown historical custom names cannot be
inferred. Name collisions between different groups still cause a skip; naming
has not been changed to resource-ID matching.

### Reviewer-friendly membership context

Group reviewer descriptions identify the immediate parent group(s) and related
application assignments discovered from the current input. For example:

> Confirm whether each user still needs access to the group Finance Readers.
> Member of 'Finance Access'. Related application assignments: 'Payroll'.

This describes recorded membership and assignment relationships, not guaranteed
effective application access. A group-only run cannot infer application
assignments that it has not discovered. Each review still covers only the
group's direct users; denying access removes that user's membership in that
group, not the group-to-group relationship.

Administrator descriptions retain every discovered membership path, including
shared parents and circular closing connections, and associate each path with
its originating application's name and ID where available. Descriptions are
captured in the result's `ReviewConfiguration`, JSONL and requested HTML report.
Historical descriptions/application names are not invented when exporting old logs.

Reviewer emails support 256 characters. Explicit `Role` and `ReviewerMessage`
text is preserved; if the full relationship names do not fit, a warning is
emitted and the reviewer description uses relationship counts instead. Full
paths remain in administrator descriptions and the report. If even that summary
cannot fit, the target fails explicitly with a request to shorten the supplied
text or group name; context and explicit role text are not silently discarded.

### Inherited settings and removal impact

Children inherit effective `ReviewerType`, `AppOwnerIds`, `Role`, and run-wide
schedule/duration/fallback settings. Conflicting effective policies for a shared
child stop the run before any creation; equivalent defaults and differently
ordered equal owner-ID sets do not conflict.

CSV `Owner` and `ReviewName` are **not inherited** from a parent. For a synced
child, provide a separate CSV row with that child's `Owner`, or configure explicit
fallback reviewers. Without either, the synced-child target fails. Cloud children
continue to use their own Entra human owners. If no human owners exist, the
existing explicit fallback rules apply.

**A denial removes membership from the child group itself**, potentially affecting
every application/resource it grants access to, not just the parent that led to
discovery. Reviews do not certify or remove the child group's membership in the
parent. Discovery of nested groups also does not imply nested group assignments
grant access to an enterprise application.

Circular membership (for example, Group A -> Group B -> Group A) produces a
warning containing group names and object IDs, not a discovery failure. The
closing connection is retained in `GroupPaths` and `ParentGroupIds` but is not
traversed again. Other branches continue normally; every distinct group still
gets at most one eligible direct-user review, with existing-name guards unchanged.
The HTML topology retains the circular arrows and displays a cycle warning.
Investigate and remediate the membership model separately; the script never
changes group-to-group membership. Cycles do not prove that nested assignments
grant effective access to an application.

Failures reading any group's metadata/membership still abort discovery before
creation, as do conflicting inherited settings. Automatically discovered
nested-only dynamic groups are explicitly skipped because membership is
rule-managed. Existing application-mode dynamic-root skip behavior is unchanged.

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
stage at every size; non-user owners are never reviewers.
Without usable user owners, the target fails. `ApplicationOwner` uses explicit
`AppOwnerIds` for the final stage at every size.

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
They are required whenever `ApplicationOwner` is selected. They are optional
at every size when a target uses
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

For a service-principal row, `Owner` applies to synchronized groups directly
assigned to that application, not their nested children or its direct-user review.
The column is
used only for synchronized groups with `ResourceOwner` routing at any size;
cloud groups use their Entra owners. A populated `Owner` takes precedence over `-FallbackReviewerUserId`
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

Groups with zero direct users are skipped with `No direct users; no access review needed`,
even when they have owners or nested groups. Nested groups are still discovered.
This applies to single-group, CSV, and application discovery.
Applications without direct users do not get a direct-user review.

| Direct users | Stages |
| --- | --- |
| 1-150 | Selected owner |
| 151-499 | Line managers -> selected owner |
| 500+ | Self-review -> line managers -> selected owner |

`-ReviewerType ApplicationOwner` (the default) ends with the supplied
`AppOwnerIds`. `-ReviewerType ResourceOwner` ends with the group's human owners,
including the third stage for 500+ members. CSV `ReviewerType` overrides the
run-wide selection per row. For example, a group with 500+ members can use:

```powershell
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -ReviewerType ResourceOwner -WhatIf
```

For new reviews, `ResourceOwner` now applies at every size; older versions
always used application owners for 500+. Existing reviews are not updated.

Use `-ForceTwoStage` to add line managers before the selected owner for targets
with 1-150 members. Targets with 151-499 members still use two stages, and 500+
retain all three stages. Empty targets remain skipped. The switch applies
run-wide in single-group, CSV, and application modes.

```powershell
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -ReviewerType ResourceOwner `
    -ForceTwoStage -WhatIf
```

Forced stages use the same owner/fallback routing and stage-duration settings
as automatic stages. An explicit one-time `EndDate` splits the date window
across the resulting stages.

`ResourceOwner` selects group user owners or, for direct-user application reviews,
service-principal user owners. Non-synchronized groups with no user owners
(including groups owned only by service principals) use explicitly configured
`-FallbackReviewerUserId` users instead; without them the target fails. If a
group has both user and service-principal owners, only its user owners review.
Direct-user application reviews without user owners still fail. For synchronized
groups at any size, CSV `Owner` users review instead of group owners.
If `Owner` is blank or absent, explicitly supplied fallback users review
instead; without either the target fails.
Group counts include distinct direct users only; direct-user
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
a warning; manager stages depend on size or `-ForceTwoStage`.
Graph advances `Approve` and `NotReviewed` decisions to subsequent stages;
denied decisions do not advance.

## Schedule and notifications

Defaults to `-Frequency OneTime` and `-StartDate` today. Recurring options:
`Weekly`, `Monthly`, `Quarterly`, `SemiAnnual`, `Annual`. Recurring `-EndDate`
defaults to the end of the current month and must allow all stages to finish.
For one-time reviews, an explicitly supplied `-EndDate` sets the total duration
from `StartDate` to `EndDate` and takes priority over stage-duration parameters.
Dates are whole-day boundaries: November 16 to November 27 is 11 days.
The window is divided evenly across the included stages, with any remaining
days assigned to earlier stages:

- One stage: 11 days.
- Two stages: 6 + 5 days.
- Three stages: 4 + 4 + 3 days.

The window must allow at least one day per stage. For example:

```powershell
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -AppOwnerIds '<application-owner user Object ID>' `
    -StartDate '2026-11-16' -EndDate '2026-11-27' -WhatIf
```

Without an explicit one-time `EndDate`, duration remains the sum of the included
stages. For recurring reviews, `EndDate` remains the series end date, not the
duration of each occurrence; supply a later date to allow subsequent occurrences.

Each stage otherwise defaults to five days. Set independent positive whole-day durations:

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
For `ResourceOwner` routing at any size, synchronized groups require CSV
owners or explicit fallback users as described above.

All modes skip an existing review with the same display name (case-insensitive),
including completed reviews; existing reviews are not updated. The script stops
without creating reviews if it cannot retrieve existing names. A preview also
skips duplicate names within the batch. Matching is by name, not resource ID, so
renaming a review or changing its prefix can allow another review for the same
resource. Concurrent runs can still create duplicates.
Application-discovery mode also skips dynamic root groups; nested-only dynamic
groups are skipped in every mode. All modes skip a group's own review when it
has no direct users, without skipping discovery of its descendants.

## Results and private logs

Results include group/review IDs, `MemberCount`, effective application-owner IDs, stage names,
reviewer type, status and failure reason. Batch summaries are written separately
from result objects. A total direct-user count for group reviews appears directly
below the batch summary and is also printed for single-group runs. It includes
only created reviews, or would-be-created reviews with `-WhatIf`; skipped,
failed, and declined reviews are excluded. The total is zero if no reviews
qualify. Users directly in more than one group count separately for each review;
this is not a tenant-wide distinct-user total. `MemberCount` now means the
distinct direct-user count used for group stage selection, excluding nested
membership and non-user objects. Older logs retain the previous transitive-count
meaning. Counts are also retained in per-group logs:

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
group direct-user totals.

All result shapes include `IsNested`, `RootGroupIds`, `ParentGroupIds`,
`SourceServicePrincipalIds`, `GroupPaths`, and `MatchedReviewName` (empty where
not applicable). `GroupPaths` retains group IDs and display names from each root
to the reviewed group, along with its originating application ID and
`ServicePrincipalName` where applicable.
Shared children retain every discovery path while producing one review.
`MatchedReviewName` identifies the canonical/legacy name causing a duplicate skip.
JSONL preserves these structured arrays; flatten them explicitly when exporting
these fields to CSV.

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
Circular paths use `GroupCycleDetected` events containing `GroupIds`,
`GroupNames`, the originating `ServicePrincipalId` where applicable, and a
diagnostic `Reason`. Paths include the repeated closing group so the cycle can
be identified. A warning/event is emitted for each encountered circular path;
the same cycle may be encountered from different roots.
CSV files themselves are not copied into logs; effective application-owner IDs,
CSV owner IDs and role text are logged in results.
Authentication commands, credentials and tokens are not logged. Keep logs
private because they contain tenant/group metadata.

WhatIf creates no log files or log directories and does not change existing logs.
It writes no files unless you explicitly request an HTML report using
`-HtmlReportPath`, as described below.
All real-run logs are kept until you remove them manually. The old combined
`logs/Create-AccessReview.jsonl` is left untouched; it is not migrated or appended to.

Open the printed path to inspect one run, or find the most recent run:

```powershell
$log = Get-ChildItem '.\logs\Create-AccessReview_*.jsonl' | Sort-Object Name -Descending | Select-Object -First 1
Get-Content -LiteralPath $log.FullName | ConvertFrom-Json
```

Logging errors stop real runs explicitly;
check IDs already returned before retrying. Single-group failures are logged
before rethrowing; batch creation/reviewer failures are reported per group.
Application/group discovery failures and shared-policy conflicts are logged as
`RunFailed` and stop before creation. Preflight errors occur before logging starts.

## Offline HTML reports

Reports are opt-in. Combine group relationships, available stage timelines and
a searchable, paginated inventory in one local HTML file:

```powershell
# Preview: writes only the requested report, not reviews or JSONL logs.
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner `
    -WhatIf -HtmlReportPath '.\reports\preview.html'

# Real run: records actual creation outcomes plus the submitted configuration.
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner `
    -HtmlReportPath '.\reports\campaign.html'

# Export an existing per-run log without Graph modules, permissions or a connection.
.\Export-AccessReviewReport.ps1 -LogPath '.\logs\Create-AccessReview_<timestamp>_<runId>.jsonl' `
    -OutputPath '.\reports\historical-campaign.html'
```

Use your actual input/reviewer settings; the examples do not change defaults.
Open the printed HTML path in a browser. Reports need no server or external
assets and make no background network calls. Nothing opens automatically.
Explicit **Open group in Entra** links navigate online in a new tab; sign in
to the correct directory with permission to view the group.
Without `-HtmlReportPath`, provisioning does not create reports or report
directories. Explicit report export can create its destination directory,
including during WhatIf. Existing output paths are rejected rather than
overwritten; use a new filename for each export.

### Views and themes

- **Summary:** recorded outcomes, with existing-name matches separated from other
  skips, and separate group/direct-application populations. Totals cover all
  recorded targets, not only created reviews; people can be counted in several
  reviews. Manual ownership/workload items remain outside user-review coverage.
- **Group topology:** visible group cards and directional parent-to-child arrows,
  including multiple connections into a shared child without duplicating it.
  Cards show recorded population, stages and outcome; select one for its details.
  Zoom, fit-to-width and scrolling keep larger diagrams navigable. A hierarchy
  selector switches between connected components or shows all groups. Above 50
  groups, disconnected campaigns initially show the first component; no groups
  or edges are discarded. The counter states how many are currently displayed.
  Referenced parents without result records are shown with unavailable population
  and stage information. Missing historical paths never imply no nesting.
- **Group links:** topology cards, review details and inventory group rows include
  **Open group in Entra**. Links are generated only from valid nonempty group
  Object IDs and open the group's overview in a separate tab. Missing/invalid
  historical IDs show an unavailable-link notice instead.
- **Hierarchy and stages:** choose a review in the topology or inventory to see parent/child
  links, all recorded paths and available stage durations, planned date boundaries,
  resolved reviewer queries/IDs, fallback queries, scope and decision progression.
  Shared groups link to the same recorded review. The view does not expand every
  subtree or thousands of timelines at once.
- **Inventory:** text search, outcome/type filters and pages of 50 targets.
  Search includes available reviewer IDs, role text, paths and diagnostics.
- **Theme:** red-and-white light styling and a dark theme. The upper-right
  **Dark mode** switch initially follows the system setting, then remembers
  your explicit choice locally. If browser/file privacy settings block storage,
  switching still works and a notice explains that it cannot remember the choice.
  Printing uses a light theme.

### What a report does and does not prove

New results/logs include `ReportSchemaVersion`, `GroupCountBasis`,
`IsSyncedGroup` and nullable `ReviewConfiguration`, consistently across result
types. Configuration captures the actual computed stage list, schedule and
request settings. A Created outcome displays **submitted configuration**;
other configured targets display **computed configuration**, not evidence of
creation. Recurring series expiry is shown separately from the first planned
instance window. Timeline dates are schedule boundaries, not inclusive day
counts or live stage dates.

An existing-name match does not fetch or verify the stored definition's stages,
dates or scope. Older logs can lack durations, resolved reviewers, sync metadata
or paths; those fields are **unavailable**, not reconstructed from names or
today's defaults. Unmarked historical `MemberCount` values may be transitive
counts and are labelled **recorded members (basis unknown)**, excluded from
verified direct-user totals.

Reports are not live dashboards or proof of completed certification/remediation.
Synced groups remain report-only; denied access must be handled in on-premises AD.
Removing a cloud child membership can affect every resource using that group.
Current forwarding of decisions is unchanged.

The offline exporter accepts one per-run JSONL file, including partial/failed
runs. It reports malformed JSON with a line number and rejects identifiable
mixed-run files. Do not feed the old combined log directly; separate it into
per-run files first. Records without run identity cannot be safely attributed
to a campaign by the exporter.

Once a run starts, a requested report is also attempted for discovery/creation
failures, labelled partial with available outcomes and diagnostics. Invalid
input that prevents a run starting does not produce an empty success report.
An export failure is explicit; if reviews were already created, check returned
IDs and JSONL before retrying. If creation and export both fail, the original
creation error is preserved and the export error is also reported.

HTML contains private tenant/group metadata and reviewer IDs, like JSONL.
Protect it before sharing. Only the theme choice is stored in browser storage.
The local `reports/` and `logs/` directories are gitignored; exports elsewhere
are your responsibility.

## Read-only status report (separate script)

[Get-AccessReviewStatusReport.ps1](./Get-AccessReviewStatusReport.ps1) reads
existing reviews without running provisioning. Keep
[AccessReview.Reporting.ps1](./AccessReview.Reporting.ps1) beside it.
It uses only Graph **GET** requests: no creation, changes, reminders or application
of decisions.

```powershell
Connect-MgGraph -Scopes 'AccessReview.Read.All','User.Read.All','GroupMember.Read.All','Application.Read.All'

$status = .\Get-AccessReviewStatusReport.ps1 `
    -Prefixes @('AR ', 'Quarterly ') `
    -HtmlReportPath '.\reports\status.html' -OpenReport
```

Use your campaign prefixes. Matching is literal, case-insensitive, and uses the
first matching prefix in the array; characters such as `[` and `*` are not
wildcards. The default is `@('AR ')`. **Every historical instance returned by
Graph** is included, not just the newest or currently running instance.
Definitions with no instances and failed instance listings remain visible.
Collection can take time for large histories because decisions are read for
each instance and runtime stage.

The script requires an existing Graph connection and the
`Microsoft.Graph.Authentication` and `Microsoft.Graph.Users` modules.
`AccessReview.Read.All` (or `AccessReview.ReadWrite.All`) permits review reads.
`User.Read.All` permits basic user identities and current manager expansion
through `Get-MgUser`. If only `User.ReadBasic.All` is granted, basic identities
can still resolve, but manager-read failures are shown as collection issues.
`GroupMember.Read.All`
supports current group owners/reviewer-group membership; `Application.Read.All`
supports current application owners. Additional permissions can be required for
hidden membership or directory restrictions. The signed-in account also needs
an appropriate Entra role for the reviews being read. App-only connections need
the corresponding application grants and consent. The connected Graph cloud
endpoint is used. Directory-read failures preserve IDs and appear as issues,
rather than being silently treated as missing owners.

### Stages, reviewers and owner identities

The red/white report includes the same upper-right dark-mode switch as creation
reports, prefix summaries, search (including names/UPNs/IDs), instance/definition
status filters and pages of 50 inventory entries. Details are loaded into the
page only when selected. The HTML has no external assets or background requests;
JavaScript enables filtering and detail selection.

For each review/instance, it separates:

- **Current configured stages:** all stages, including future stages not yet
  returned by the runtime API, with duration, dependencies, reviewer rules,
  resolved named users and fallbacks.
- **Runtime stages:** Graph's stage GUIDs, statuses, dates, reviewer/fallback
  scopes, decision counts, individual subjects (including pending self-reviews),
  recorded decision-makers, decision timestamps and application results. Runtime stages are
  sorted for display, **not guessed into a mapping to current configuration**.
- **Recorded instance reviewers:** Graph's `contactedReviewers`, retaining
  recorded names and UPNs where present. These are instance-wide, not a per-stage
  assignment roster or proof of notification delivery.
- **Current resource owners:** human owners with names, UPNs and IDs, separated
  from service-principal/unclassified owners. These are **today's directory
  owners**, not historical stage assignments. Group resources include safe
  Entra links.
- **Individual users and current line managers:** instance and stage decision
  tables identify each user being reviewed and the recorded decision-maker.
  When the current definition contains a manager-review stage, each user also
  has a separately labelled current manager (or an explicit unavailable/no-manager
  result). These current relationships are **not historical assignments**.
  Pending items and automatic/default decisions can have no recorded actor;
  that does not mean no reviewer was assigned. Profiles and manager lookups are
  cached for the run, including failures.

For reviews created by this tool, resource-owner policies are resolved to
explicit `/users/{id}` reviewer queries when created. Graph does not preserve
an original “resource owner” policy label on those IDs. The status report
resolves and shows these configured users, but cannot distinguish captured
owners from manually selected users or infer an earlier owner roster.
Current profiles may have changed; the report labels their evidence source.
Named reviewer queries in `/users/{id}`, `/v1.0/users/{id}` and absolute URLs
on the connected Graph endpoint are resolved to names and sign-in addresses.
Technical queries/JSON are collapsed rather than used as the primary identity
display. Unrecognized queries are retained without following arbitrary URLs.

Simple dynamic owner rules are expanded to current human owners. Simple
reviewer-group membership rules are expanded to current human members.
Per-person manager/self-review rules cannot supply a complete historical
per-stage roster from configuration alone; consult recorded instance reviewers
and individual stage decisions. Current managers are supporting directory
context, not a reconstruction of past assignments. Filtered/unsupported rules remain visible verbatim,
without claiming the separate unfiltered current owner list is eligible.
Raw resource/reviewer scopes are available in details.

### Status interpretation and failures

Overview and prefix decision totals use **instance-level** decision items only;
stage tasks are not added again. People may appear in multiple instances or
reviews. Approve, Deny and Don't know are recorded outcomes, not necessarily
human actions: automation/default decisions can produce them. Unknown values
(including NotEnabled) are shown as **Other**, not completed. No decision items
means **n/a**, not 100%. Application-result breakdowns are separate: successful
application is not proof of denied-access removal.

Decision-count coverage states how many instance summaries were available.
A failed page invalidates that summary rather than presenting partial counts
as complete. Stage/contact/owner/profile failures produce warnings and visible
issues without erasing other useful results. A global collection failure
attempts a clearly partial HTML export, then rethrows the original error.
Missing runtime stages never silently become a single-stage review.
This is a non-transactional snapshot of what Graph exposes, not a live dashboard
or proof of certification/remediation.

Without `-HtmlReportPath`, a unique file is created under `reports/`. Existing
output paths are rejected. Only `-OpenReport` opens a browser. The returned
`$status` object contains structured definitions, instances, stages, summaries,
owners and issues for further inspection. HTML and returned data contain
private user identities: protect them before sharing.

## Contributor and agent guidance

This tool already belongs to the Mumsbitar Git repository; it needs no separate
repository. Tool-specific contributor/agent guidance is in
[AGENTS.md](./AGENTS.md), with path-scoped Copilot guidance in
[access-reviews.instructions.md](../../../.github/instructions/access-reviews.instructions.md).
The scoped file avoids applying access-review policy to unrelated tools;
`.github/copilot-instructions.md` would instead be repository-wide.

## Offline validation

```powershell
Invoke-Pester -Path '.\tests' -Output Detailed
```

Tests mock Graph; they never connect to or change a tenant.
