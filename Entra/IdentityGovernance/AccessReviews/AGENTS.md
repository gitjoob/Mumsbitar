# AccessReviews contributor guidance

This guidance applies to this tool and its tests, not the rest of the repository.

## Entry points

- [Create-AccessReview.ps1](./Create-AccessReview.ps1): Graph discovery, policy and provisioning.
- [AccessReview.Logging.ps1](./AccessReview.Logging.ps1): private per-run JSONL events.
- [AccessReview.Reporting.ps1](./AccessReview.Reporting.ps1): shared offline HTML renderer.
- [Export-AccessReviewReport.ps1](./Export-AccessReviewReport.ps1): offline per-run JSONL export; no Graph dependency.
- [README.md](./README.md): supported inputs, policy, permissions and examples.

## Preserve these invariants

- Keep existing options and defaults unless their change is explicitly requested.
- Stage thresholds count distinct direct users; discover nested groups separately,
  merge by group ID and retain all paths, including circular closing connections.
  Warn/log cycles and stop only that recursive branch; metadata/membership read
  failures and policy conflicts still fail closed before creation.
- Preserve case-insensitive existing-name guards and aliases. Do not repair,
  stop, delete or recreate existing/running reviews as a side effect of a rerun.
- Resource-owner routing uses each group's human owners. Synced children do not
  inherit parent CSV Owner; use group-specific Owner or explicit fallback users.
- Preserve ForceTwoStage, selected final owner, explicit one-time date-window
  division, recurring-series semantics and configured decision progression.
- Keep tenant writes behind ShouldProcess. WhatIf writes neither reviews nor
  logs; explicitly supplying HtmlReportPath permits only the requested report.
- Reports describe computed/submitted settings and recorded outcomes, not live
  progress/remediation. Never infer unavailable historical configuration or
  treat an unmarked old member count as a direct-user count.
- Keep logs/reports private, HTML-encode untrusted metadata, and keep reporting
  offline without external assets or new Graph queries.
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
