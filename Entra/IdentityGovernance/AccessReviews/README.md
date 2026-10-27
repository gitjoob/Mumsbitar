# Access review scripts

These scripts create an access review for a group's direct members. They default
to a one-time review. Set `-Frequency` to `Weekly`, `Monthly`, `Quarterly`,
`SemiAnnual`, or `Annual` for a recurring review. `-StartDate` defaults to today.
For recurring reviews, `-EndDate` defaults to the end of the current month.
`-ReviewName` is optional; otherwise, the group name is used.

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

In both scripts, `$AdditionalNotificationRecipients` optionally lists user or
group IDs for a completion notification; leave it empty to omit that
notification. Unanswered reviews default to Deny, and denied access is
automatically removed. Use `-WhatIf` to preview the operation.

Both scripts require `-GroupId`. For example:

```powershell
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -WhatIf
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -Frequency Quarterly -WhatIf
```
