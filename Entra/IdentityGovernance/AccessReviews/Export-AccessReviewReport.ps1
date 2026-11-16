<#
.SYNOPSIS
Exports one access-review run's JSONL log to a private, offline HTML report.
.DESCRIPTION
Does not import Graph or contact Entra. Existing output files are never overwritten.
Older logs retain their recorded values; missing configuration and count basis
are explicitly unavailable. Mixed-run or malformed logs fail explicitly.
.EXAMPLE
.\Export-AccessReviewReport.ps1 -LogPath '.\logs\Create-AccessReview_<timestamp>_<runId>.jsonl' -OutputPath '.\reports\campaign.html'
#>
[CmdletBinding()]
param(
	[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LogPath,
	[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputPath
)

. (Join-Path $PSScriptRoot 'AccessReview.Reporting.ps1')
$OutputPath = Resolve-AccessReviewReportPath -Path $OutputPath
$entries = [System.Collections.Generic.List[object]]::new()
$lineNumber = 0
foreach ($line in Get-Content -LiteralPath $LogPath -ErrorAction Stop) {
	$lineNumber++
	if ([string]::IsNullOrWhiteSpace($line)) { continue }
	try {
		$entry = ConvertFrom-Json -InputObject $line -ErrorAction Stop
		if (-not $entry -or -not $entry.Event -or $entry -is [array]) { throw 'Expected an event object.' }
		$entries.Add($entry)
	}
	catch { throw "Invalid JSONL event in '$LogPath' at line ${lineNumber}: $($_.Exception.Message)" }
}
$runIds = @($entries | ForEach-Object { $_.RunId } | Where-Object { $_ } | Sort-Object -Unique)
if ($runIds.Count -gt 1 -or @($entries | Where-Object Event -eq 'RunStarted').Count -gt 1 -or
	@($entries | Where-Object Event -eq 'RunCompleted').Count -gt 1) {
	throw 'The log contains multiple runs. Export a separate per-run JSONL file; campaigns are never combined automatically.'
}
if (-not @($entries | Where-Object { $_.Event -in @('RunStarted', 'RunFailed', 'GroupResult', 'ApplicationResult', 'ManualReviewRequired') }).Count) {
	throw 'No supported access-review run events were found in the log.'
}
$results = @($entries | Where-Object { $_.Event -in @('GroupResult', 'ApplicationResult', 'ManualReviewRequired') } | ForEach-Object {
	if (-not $_.Result) { throw "Log event '$($_.Event)' is missing its Result object." }
	$_.Result
})
$context = @{
	Mode = 'Offline JSONL export (recorded outcomes, not live state)'
	RunId = if ($runIds.Count) { $runIds[0] } else { $null }
	Completed = @($entries | Where-Object Event -eq 'RunCompleted').Count -gt 0
	Failure = (@($entries | Where-Object Event -eq 'RunFailed' | ForEach-Object { $_.Reason }) -join '; ')
}
Write-AccessReviewReport -Path $OutputPath -Results $results -Context $context
