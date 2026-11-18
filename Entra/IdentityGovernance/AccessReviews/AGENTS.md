# AccessReviews contributor guidance

This guidance applies to this tool and its tests, not the rest of the repository.

## Entry points

- [Create-AccessReview.ps1](./Create-AccessReview.ps1): Graph discovery, policy and provisioning.
- [AccessReview.Logging.ps1](./AccessReview.Logging.ps1): private per-run JSONL events.
- [AccessReview.Reporting.ps1](./AccessReview.Reporting.ps1): shared offline HTML renderer.
- [Export-AccessReviewReport.ps1](./Export-AccessReviewReport.ps1): offline per-run JSONL export; no Graph dependency.
- [Get-AccessReviewStatusReport.ps1](./Get-AccessReviewStatusReport.ps1): separate,
  read-only Graph status collection covering all historical instances.
- [README.md](./README.md): supported inputs, policy, permissions and examples.

## Preserve these invariants

- Keep existing options and defaults unless their change is explicitly requested.
- Stage thresholds count distinct direct users; discover nested groups separately,
  merge by group ID and retain all paths, including circular closing connections.
  Warn/log cycles and stop only that recursive branch; metadata/membership read
  failures and policy conflicts still fail closed before creation.
- Preserve case-insensitive existing-name guards and aliases. Do not repair,
  stop, delete or recreate existing/running reviews as a side effect of a rerun.
- Also guard by resource ID/scope across prefixes for active/scheduled and
  potentially future recurring coverage, irrespective of reviewer policy.
  Completed historical coverage alone must not block by resource. Partial or
  unresolvable overlapping scopes require explicit manual assessment, not an
  assumption of full coverage. Keep group membership and application assignment
  scopes distinct; a transitive parent review does not certify child membership.
- New nested-only names omit `[Nested]`; preserve that suffix as a legacy alias.
  Explain recorded parent/application relationships in descriptions without
  claiming nested assignments prove effective application access. Preserve
  explicit Role/ReviewerMessage text within the reviewer email limit.
- Resource-owner routing uses each group's human owners. Synced children do not
  inherit parent CSV Owner; use group-specific Owner or explicit fallback users.
- Preserve ForceTwoStage, selected final owner, explicit one-time date-window
  division, recurring-series semantics and configured decision progression.
- Keep tenant writes behind ShouldProcess. WhatIf writes neither reviews nor
  logs; explicitly supplying HtmlReportPath permits only the requested report.
- Creation/offline reports describe computed/submitted settings and recorded
  outcomes, not live progress/remediation. Never infer unavailable historical
  configuration or treat an unmarked old member count as a direct-user count.
- Status collection uses GET only. Keep current configured stages, runtime
  stages, recorded instance reviewers, decision actors and current directory
  owners separate. Do not infer the original owner-policy label from named
  user scopes or map runtime stage GUIDs to current configuration by position.
  Count instance decisions once, never add stage totals to overview completion,
  and distinguish unavailable data from zero. Retain individual decision
  subjects and actors, including pending subjects. Resolve user profiles with
  Get-MgUser; keep current managers explicitly separate from historical
  assignments and avoid treating absent decision actors as absent reviewers.
- Keep logs/reports private and HTML-encode untrusted metadata. Renderers and
  creation/offline reporting make no new Graph queries; only the separate status
  collector reads Graph. Generated HTML stays offline without external assets.
- Surface failures explicitly; report-export errors must not obscure the original
  provisioning error or encourage blind retries after successful creation.

## Validation

Run from this directory with PowerShell and the existing Graph/Pester modules:

```powershell
$r = Invoke-Pester -Path ./tests -Output Detailed -PassThru
if ($r.FailedCount) { exit 1 }
```

Tests mock Graph. Never connect to a tenant or create/modify real reviews for
validation. Add focused behavioral tests in the existing test files and run
`git diff --check`. Browser-check report interactions with fictional data.
Do not install dependencies unless the chosen validation reports them missing.
