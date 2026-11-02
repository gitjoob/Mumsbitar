# Access review scripts

These scripts create an access review for a group's direct members. They default
to a one-time review. Set `-Frequency` to `Weekly`, `Monthly`, `Quarterly`,
`SemiAnnual`, or `Annual` for a recurring review. `-StartDate` defaults to today.
For recurring reviews, `-EndDate` defaults to the end of the current month;
it is ignored for one-time reviews, whose period is `-StartDate` plus the days
of all stages (a warning is shown if you pass it).
`-ReviewName` is optional; otherwise, the group name is used.

## Reviewer selection and creation logs

All three scripts accept `-ReviewerType ResourceOwner` (group owners) or
`-ReviewerType LineManager` (each reviewee's manager). Existing defaults and
stage layouts are preserved:

| Script | Default | What the selector changes |
| --- | --- | --- |
| `Create-AccessReview.ps1` | `ApplicationOwner` | Reviewers of the single stage. `ApplicationOwner` remains an additional supported value for the configured reviewers. |
| `Create-AccessReviewStages.ps1` | `LineManager` | The stage before application owners in two-/three-stage reviews. Single-stage reviews still use application owners. |
| `Create-AccessReviewAppGroups.ps1` | `ResourceOwner` | The final reviewer stage; an initial self-review stage is unchanged. |

Missing managers or group user owners use the configured fallback reviewer.
Owner lookup errors fail explicitly. The application-group script retains its
skip rules for dynamic groups, groups with no members and no user owners, and
existing review names.

Each script appends invocation syntax, timestamps, run ID, effective schedule,
reviewer selection, per-group results and created review IDs to its own
`logs/<script-name>.jsonl`. The whole `logs/` directory is gitignored.
`-WhatIf` writes local preview logs, never Graph reviews. Logging errors stop
the run; inspect created IDs before retrying. Single-group failures are logged
before being rethrown, and skipped application groups are logged as well.
Preflight validation failures happen before logging starts.

Keep the shared [AccessReview.Logging.ps1](./AccessReview.Logging.ps1) beside
all three scripts when copying them. It logs explicit invocation parameters,
not the authentication command, credentials or tokens. Logs contain group and
tenant metadata and must be kept private. The logged working directory is
needed when replaying commands with relative CSV paths.

```powershell
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner -WhatIf
.\Create-AccessReviewAppGroups.ps1 -ServicePrincipalId '00000000-0000-0000-0000-000000000000' -ReviewerType LineManager -WhatIf
Invoke-Pester -Path '.\tests'
```

## Scripts

- [Create-AccessReview.ps1](./Create-AccessReview.ps1) creates a single-stage
  review. Reviewers are configured in `$ApplicationOwners`; the fallback
  reviewer is configured in `$FallbackReviewerUserId`.
- [Create-AccessReviewStages.ps1](./Create-AccessReviewStages.ps1) selects one,
  two, or three stages based on the group's transitive member count. Configure
  the ascending cutoffs in `$StageThresholds`: below the second cutoff uses
  application owners; at the second, line managers then application owners;
  at the third, self-review, line managers, then application owners. Counts
  below the first cutoff still use one stage and produce a warning.

  Set `$ApplicationOwners` to one or more user or group IDs for the application
  owner stage; group membership is resolved transitively. `$FallbackReviewerUserId`
  is used when a primary reviewer is unavailable, and `$StageDurationInDays`
  sets the duration of each stage. The default thresholds and placeholder IDs
  are defined at the top of the script.

  Use `-ReviewerType LineManager` (default) or `-ReviewerType ResourceOwner`
  to choose the reviewer stage before application owners in two-/three-stage
  reviews. Resource owners are the group's user owners (service principals are
  excluded); groups without user owners use `$FallbackReviewerUserId`. Owner
  lookup errors fail the group instead of silently using the fallback.
  Single-stage reviews still use application owners; the final stage and
  optional initial self-review stage are unchanged.

  Validated runs append JSON Lines to
  `logs/Create-AccessReviewStages.jsonl` beside the script. This directory is
  gitignored. Entries contain a run ID, UTC timestamp, runnable invocation
  syntax with explicitly supplied parameters, working directory, effective
  dates/frequency, reviewer selection, per-group status, created review ID and
  any failure reason. `RunCompleted` contains totals. Single-group failures
  are logged before being rethrown; preflight failures occur before logging
  starts. `-WhatIf` writes local preview logs but never creates Graph reviews.
  Logging failures stop the script explicitly; check existing review IDs
  before retrying. Logs contain tenant/group metadata, so keep them private.

  ```powershell
  .\Create-AccessReviewStages.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner -WhatIf
  .\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -ReviewerType LineManager
  Get-Content '.\logs\Create-AccessReviewStages.jsonl' | ForEach-Object { $_ | ConvertFrom-Json }
  ```

  Offline reviewer-selection and logging tests:
  `Invoke-Pester -Path '.\tests\Create-AccessReviewStages.Tests.ps1'`.
- [Create-AccessReviewAppGroups.ps1](./Create-AccessReviewAppGroups.ps1) takes
  `-ServicePrincipalId` instead of `-GroupId`. An access review of an
  application can't use group owners as reviewers, and it doesn't remove users
  who get access through a group. This script therefore creates one review per
  group assigned to the enterprise application, and each group's owners review
  their own group. Groups with more than `$SelfReviewStageThreshold` transitive
  members (default 150) get a self-review stage first. Groups without
  owners go to `$FallbackReviewerUserId`. Groups that already have a review with
  the same name (`AR <application> - <group>`) are skipped. Each group gets a result with its status. At the end, the
  script lists only the skipped or failed groups with the reason, and the number
  of reviews created (with `-WhatIf`, the number that would be created).

In all scripts, `$ReviewerMessage` optionally adds your own text (for example
contact details or a link) to the reviewer email. The email supports up to 256
characters in total, including the default text; the script stops if it is too
long. The rest of the email is fixed by Microsoft.

In all scripts, `$AdditionalNotificationRecipients` optionally lists user or
group IDs for a completion notification; leave it empty to omit that
notification. Unanswered reviews default to Deny, and denied access is
automatically removed. Use `-WhatIf` to preview the operation.

If the group is synchronized from on-premises Active Directory
(`onPremisesSyncEnabled`), the access review cannot change its membership. The
scripts detect this, disable automatic apply, and show a warning; decisions are
only recorded and denied access must be removed in Active Directory. Download
the CSV report from the completed review, or retrieve the decisions through
Microsoft Graph, and act on them in Active Directory. Synced groups have no
owners in Microsoft Entra ID, so the fallback reviewer is used. The
application script skips groups with dynamic membership and groups with
no members and no owners.

`Create-AccessReview.ps1` and `Create-AccessReviewStages.ps1` also accept
`-CsvPath` instead of `-GroupId` to create one review per group. The CSV needs
a `GroupId` column and can have an optional `ReviewName` column; duplicate
groups are processed once. A failure for one group doesn't stop the others;
failed groups and a total are listed at the end.

```csv
GroupId,ReviewName
00000000-0000-0000-0000-000000000000,
00000000-0000-0000-0000-000000000000,Quarterly review of Finance
```

The group scripts require `-GroupId` or `-CsvPath`. For example:

```powershell
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -WhatIf
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -WhatIf
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -Frequency Quarterly -WhatIf
.\Create-AccessReviewAppGroups.ps1 -ServicePrincipalId '00000000-0000-0000-0000-000000000000' -WhatIf
```
