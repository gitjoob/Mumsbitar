function Resolve-AccessReviewReportPath {
	param([Parameter(Mandatory)][string]$Path)

	$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
	if ([System.IO.Path]::GetExtension($fullPath) -notin @('.html', '.htm')) {
		throw 'The HTML report output path must end in .html or .htm.'
	}
	if (Test-Path -LiteralPath $fullPath) {
		throw "Report output already exists: '$fullPath'. Choose a new filename; reports are never overwritten."
	}
	$parent = [System.IO.Path]::GetDirectoryName($fullPath)
	while ($parent -and -not [System.IO.Directory]::Exists($parent)) {
		if ([System.IO.File]::Exists($parent)) { throw "Report output directory is blocked by a file: '$parent'." }
		$parent = [System.IO.Path]::GetDirectoryName($parent)
	}
	$fullPath
}

function ConvertTo-AccessReviewHtmlText {
	param([AllowNull()][object]$Value)
	[System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-AccessReviewReportOutcome {
	param([object]$Result)
	if ($Result.MatchedReviewName) { 'Existing name match' }
	elseif ($Result.Status) { [string]$Result.Status }
	else { 'Unknown outcome' }
}

function New-AccessReviewGroupLink {
	param([AllowNull()][string]$GroupId)

	$id = [guid]::Empty
	if ([guid]::TryParse($GroupId, [ref]$id) -and $id -ne [guid]::Empty) {
		$url = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/$($id.ToString())"
		"<a class=`"entra-link`" href=`"$url`" target=`"_blank`" rel=`"noopener noreferrer`">Open group in Entra</a>"
	}
	else {
		'<small class="muted">Entra link unavailable: valid group Object ID not recorded.</small>'
	}
}

function New-AccessReviewHtml {
	param(
		[AllowEmptyCollection()][object[]]$Results = @(),
		[Parameter(Mandatory)][hashtable]$Context
	)

	$body = [System.Text.StringBuilder]::new()
	$null = $body.AppendLine('<main><header><div><p class="eyebrow">Identity governance / local report</p><h1>Access review campaign</h1></div><div class="theme-control"><button id="theme" type="button" role="switch" aria-checked="false" aria-label="Dark mode">Dark mode</button><small id="theme-notice" role="status"></small></div></header>')
	$mode = ConvertTo-AccessReviewHtmlText $Context.Mode
	$runId = ConvertTo-AccessReviewHtmlText $Context.RunId
	$null = $body.AppendLine("<p>$mode | Run: $(if ($runId) { $runId } else { 'not recorded' })</p>")
	$null = $body.AppendLine('<p class="notice">Configuration and recorded creation outcomes only. This is not live review progress, proof of certification, or proof that denied access was removed. Names and reviewer IDs in this file are private.</p>')
	if ($Context.Failure -or -not $Context.Completed) {
		$failure = ConvertTo-AccessReviewHtmlText $Context.Failure
		$null = $body.AppendLine("<p class=`"notice warning`" role=`"alert`"><strong>Partial or failed run.</strong> $failure No completion is inferred for missing targets.</p>")
	}
	$outcomes = @{}
	foreach ($result in $Results) {
		$outcome = Get-AccessReviewReportOutcome $result
		if (-not $outcomes.ContainsKey($outcome)) { $outcomes[$outcome] = 0 }
		$outcomes[$outcome]++
	}
	$null = $body.AppendLine('<section aria-label="Run summary" class="metrics">')
	foreach ($status in @('Created', 'Would be created', 'Existing name match', 'Failed', 'Skipped', 'Not created', 'Manual review required', 'Unknown outcome')) {
		$count = if ($outcomes.ContainsKey($status)) { $outcomes[$status] } else { 0 }
		$null = $body.AppendLine("<div class=`"metric`"><strong>$count</strong><span>$(ConvertTo-AccessReviewHtmlText $status)</span></div>")
	}
	$null = $body.AppendLine('</section>')
	$groups = @($Results | Where-Object { $_.GroupId -and $_.TargetType -notin @('ApplicationDirectUsers', 'ServicePrincipalOwner', 'ApplicationNonUserAssignment') })
	$knownGroups = @($groups | Where-Object { $_.GroupCountBasis -eq 'DirectUsers' -and $null -ne $_.MemberCount })
	$unknownGroups = @($groups | Where-Object { $_.GroupCountBasis -ne 'DirectUsers' -and $null -ne $_.MemberCount })
	$groupTotal = [long](($knownGroups | Measure-Object MemberCount -Sum).Sum)
	$appTotal = [long](($Results | Where-Object TargetType -eq 'ApplicationDirectUsers' | Measure-Object SubjectCount -Sum).Sum)
	$null = $body.AppendLine("<p class=`"muted`">All recorded targets: <strong>$groupTotal</strong> verified direct-user counts summed across group targets; <strong>$appTotal</strong> direct-application user counts, separately. These are per-review populations, not unique people. <strong>$($unknownGroups.Count)</strong> group count(s) have an unknown historical basis and are excluded from the direct-user total.</p>")
	$null = $body.AppendLine('<nav><a href="#topology">Group topology</a><a href="#configuration">Stages and details</a><a href="#inventory">Search inventory</a></nav>')
	$groupIndex = @{}
	$parents = @{}
	$children = @{}
	$pathLabels = @{}
	$incompleteTopology = $false
	for ($i = 0; $i -lt $Results.Count; $i++) {
		$result = $Results[$i]
		if (-not $result.GroupId) { continue }
		$id = [string]$result.GroupId
		if (-not $groupIndex.ContainsKey($id)) { $groupIndex[$id] = $i }
		if (-not $parents.ContainsKey($id)) { $parents[$id] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase) }
		foreach ($parentId in @($result.ParentGroupIds)) {
			if ($parentId) { $null = $parents[$id].Add([string]$parentId) }
		}
		foreach ($path in @($result.GroupPaths)) {
			$ids = @($path.GroupIds)
			$names = @($path.GroupNames)
			for ($p = 0; $p -lt $ids.Count; $p++) {
				if ($ids[$p] -and $p -lt $names.Count -and $names[$p]) { $pathLabels[[string]$ids[$p]] = $names[$p] }
			}
			for ($p = 1; $p -lt $ids.Count; $p++) {
				if (-not $ids[$p] -or -not $ids[$p - 1]) {
					$incompleteTopology = $true
					continue
				}
				$childId = [string]$ids[$p]
				if (-not $parents.ContainsKey($childId)) { $parents[$childId] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase) }
				$null = $parents[$childId].Add([string]$ids[$p - 1])
			}
		}
	}
	foreach ($childId in $parents.Keys) {
		foreach ($parentId in $parents[$childId]) {
			if (-not $children.ContainsKey($parentId)) { $children[$parentId] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase) }
			$null = $children[$parentId].Add($childId)
		}
	}
	$topologyIds = @(@($groupIndex.Keys) + @($parents.Keys) + @($children.Keys) | Sort-Object -Unique)
	$topologyIndex = @{}
	for ($i = 0; $i -lt $topologyIds.Count; $i++) { $topologyIndex[$topologyIds[$i]] = $i }
	$null = $body.AppendLine('<section id="topology" class="panel"><h2>Group topology</h2><p>Arrows show recorded parent-to-child membership. Shared children appear once, with a connection from every recorded parent. Choose a group card for review details, or open the group in Entra in a new tab.</p><div class="topology-controls"><label>Connected hierarchy<select id="topology-component" aria-controls="topology-viewport"><option value="all">All recorded groups</option></select></label><div><button id="zoom-out" type="button" aria-label="Zoom topology out">-</button> <button id="zoom-in" type="button" aria-label="Zoom topology in">+</button> <button id="zoom-fit" type="button">Fit width</button> <button id="zoom-reset" type="button">100%</button></div></div><p id="topology-count" class="muted" role="status"></p><p id="topology-warning" class="notice warning" role="status" hidden></p>')
	if ($incompleteTopology) { $null = $body.AppendLine('<p class="notice warning">Some recorded path IDs are missing; those incomplete connections cannot be drawn. Other recorded nodes and connections remain visible.</p>') }
	$null = $body.AppendLine('<div id="topology-viewport" tabindex="0" aria-label="Scrollable group topology"><div id="topology-space"><div id="topology-canvas"><svg id="topology-lines" aria-hidden="true"><defs><marker id="topology-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z"/></marker></defs>')
	foreach ($childId in @($parents.Keys | Sort-Object)) {
		foreach ($parentId in @($parents[$childId] | Sort-Object)) {
			$null = $body.AppendLine("<path class=`"topology-edge`" data-from=`"node-$($topologyIndex[$parentId])`" data-to=`"node-$($topologyIndex[$childId])`" marker-end=`"url(#topology-arrow)`"/>")
		}
	}
	$null = $body.AppendLine('</svg>')
	foreach ($id in $topologyIds) {
		$index = $topologyIndex[$id]
		$parentCount = @($parents[$id] | Where-Object { $_ }).Count
		$badge = if ($parentCount -gt 1) { "Shared child / $parentCount parents" } elseif ($parentCount -eq 1) { 'Nested group' } else { 'Root / no parent recorded' }
		$null = $body.AppendLine("<article id=`"node-$index`" class=`"topology-node`"><small class=`"eyebrow`">$badge</small>")
		if ($groupIndex.ContainsKey($id)) {
			$resultIndex = $groupIndex[$id]
			$result = $Results[$resultIndex]
			$name = if ($result.Group) { $result.Group } elseif ($result.Target) { $result.Target } else { $result.ReviewName }
			$population = if ($null -eq $result.MemberCount) { 'Population unavailable' }
				elseif ($result.GroupCountBasis -eq 'DirectUsers') { "$($result.MemberCount) direct users" }
				else { "$($result.MemberCount) recorded members; basis unknown" }
			$stages = if ($result.StageNames) { $result.StageNames } else { 'Stage configuration unavailable' }
			$null = $body.AppendLine("<button class=`"link view topology-name`" data-target=`"review-$resultIndex`">$(ConvertTo-AccessReviewHtmlText $name)</button><span>$(ConvertTo-AccessReviewHtmlText $population)</span><small class=`"topology-stages`" title=`"$(ConvertTo-AccessReviewHtmlText $stages)`">$(ConvertTo-AccessReviewHtmlText $stages)</small><small>$(ConvertTo-AccessReviewHtmlText (Get-AccessReviewReportOutcome $result))$(if ($result.IsSyncedGroup -eq $true) { ' / synced, report only' })</small>")
		}
		else {
			$name = if ($pathLabels.ContainsKey($id)) { $pathLabels[$id] } else { $id }
			$null = $body.AppendLine("<strong class=`"topology-name`">$(ConvertTo-AccessReviewHtmlText $name)</strong><small>Referenced in discovery paths; no review result recorded. Population and stages unavailable.</small>")
		}
		$null = $body.AppendLine("$(New-AccessReviewGroupLink -GroupId $id)</article>")
	}
	$null = $body.AppendLine('</div></div></div><noscript><p>JavaScript is required to lay out the connected topology. Group cards remain listed without connectors.</p></noscript><p class="muted">Connections reflect recorded discovery, not live membership. Groups without recorded paths are not assumed to be unrelated. Entra links require signing in to the appropriate directory with permission to view the group.</p></section><section id="configuration" class="panel"><h2>Hierarchy and stage configuration</h2><p id="selection-hint" class="muted">Choose a group in the topology or a review in the inventory to see its relationships, paths and available configuration.</p>')
	for ($i = 0; $i -lt $Results.Count; $i++) {
		$result = $Results[$i]
		$name = if ($result.ReviewName) { $result.ReviewName } elseif ($result.Target) { $result.Target } else { 'Unnamed recorded item' }
		$null = $body.AppendLine("<article id=`"review-$i`" class=`"review-detail`" hidden><h3 tabindex=`"-1`">$(ConvertTo-AccessReviewHtmlText $name)</h3><p>$(ConvertTo-AccessReviewHtmlText (Get-AccessReviewReportOutcome $result))</p>")
		if ($result.GroupId) {
			$id = [string]$result.GroupId
			$null = $body.AppendLine("<p>$(New-AccessReviewGroupLink -GroupId $id)</p>")
			$null = $body.AppendLine('<div class="relationships">')
			foreach ($direction in @('Parents', 'Children')) {
				if ($direction -eq 'Children') {
					$population = if ($null -eq $result.MemberCount) { 'Population unavailable' }
						elseif ($result.GroupCountBasis -eq 'DirectUsers') { "$($result.MemberCount) direct users" }
						else { "$($result.MemberCount) recorded members; basis unknown" }
					$null = $body.AppendLine("<div class=`"selected-group`"><small>Selected group / direct-membership scope</small><strong>$(ConvertTo-AccessReviewHtmlText $name)</strong><span>$(ConvertTo-AccessReviewHtmlText $population)</span></div>")
				}
				$related = if ($direction -eq 'Parents') { @($parents[$id]) } else { @($children[$id]) }
				$null = $body.AppendLine("<div><strong>$direction</strong><ul>")
				foreach ($relatedId in @($related | Where-Object { $_ } | Sort-Object)) {
					if ($groupIndex.ContainsKey($relatedId)) {
						$index = $groupIndex[$relatedId]
						$label = if ($Results[$index].Group) { $Results[$index].Group } else { $relatedId }
						$null = $body.AppendLine("<li><button class=`"link view`" data-target=`"review-$index`">$(ConvertTo-AccessReviewHtmlText $label)</button></li>")
					}
					else { $null = $body.AppendLine("<li>$(ConvertTo-AccessReviewHtmlText $relatedId) (no result recorded)</li>") }
				}
				if (-not @($related | Where-Object { $_ }).Count) {
					$null = $body.AppendLine('<li>None recorded; historical relationships may be unavailable.</li>')
				}
				$null = $body.AppendLine('</ul></div>')
			}
			$null = $body.AppendLine('</div><p class="notice">A child review covers its direct users, not its membership in the parent. Removal can affect every resource using the group. Nested groups do not necessarily convey enterprise-application assignments.</p>')
		}
		$config = $result.ReviewConfiguration
		if ($config -and @($config.Stages).Count) {
			$label = if ($result.Status -eq 'Created') { 'Submitted configuration (not fetched live)' } else { 'Computed configuration (not evidence of creation)' }
			$null = $body.AppendLine("<h4>$label</h4><p>Frequency: $(ConvertTo-AccessReviewHtmlText $config.Frequency). First planned instance: $(ConvertTo-AccessReviewHtmlText $config.StartDate), $([int]$config.InstanceDurationInDays) days.</p>")
			if ($config.SeriesEndDate) {
				$null = $body.AppendLine("<p>Recurring series expires: $(ConvertTo-AccessReviewHtmlText $config.SeriesEndDate). This is not an individual instance's end date.</p>")
			}
			$handling = if ($config.AutoApplyDecisionsEnabled -eq $true) { 'Automatic application configured; this is not proof of removal.' }
				elseif ($config.AutoApplyDecisionsEnabled -eq $false) { 'Report only. Remediate denied access in on-premises AD for synced groups.' }
				else { 'Application behavior unavailable.' }
			$null = $body.AppendLine("<p class=`"notice`">$handling Default decision: $(ConvertTo-AccessReviewHtmlText $config.DefaultDecision).</p><div class=`"timeline`">")
			foreach ($stage in $config.Stages) {
				$days = [int]$stage.DurationInDays
				if ($days -lt 1) { throw 'Recorded stage durations must be positive whole days.' }
				$null = $body.AppendLine("<div class=`"stage`" style=`"flex-grow:$days`"><strong>$(ConvertTo-AccessReviewHtmlText $stage.Name)</strong><span>$days days</span><small>$(ConvertTo-AccessReviewHtmlText $stage.StartDate) &rarr; $(ConvertTo-AccessReviewHtmlText $stage.EndDate)</small></div>")
			}
			$null = $body.AppendLine('</div><p class="muted">Dates are planned schedule boundaries, not inclusive day counts or live stage dates. Progression is conditional, not unanimous approval.</p><details><summary>Reviewer, scope and progression details</summary>')
			foreach ($stage in $config.Stages) {
				$reviewers = @($stage.Reviewers | ForEach-Object { $_.query }) -join ', '
				if (-not $reviewers) { $reviewers = if ($stage.Name -eq 'Self-review') { 'Each user (self-review)' } else { 'Not recorded' } }
				$fallbacks = @($stage.FallbackReviewers | ForEach-Object { $_.query }) -join ', '
				$moves = @($stage.DecisionsThatWillMoveToNextStage) -join ', '
				$null = $body.AppendLine("<p><strong>$(ConvertTo-AccessReviewHtmlText $stage.Name)</strong><br>Reviewers: $(ConvertTo-AccessReviewHtmlText $reviewers)<br>Fallbacks: $(if ($fallbacks) { ConvertTo-AccessReviewHtmlText $fallbacks } else { 'None recorded' })<br>Decisions forwarded: $(if ($moves) { ConvertTo-AccessReviewHtmlText $moves } else { 'None / final stage' })</p>")
			}
			$null = $body.AppendLine("<pre>$(ConvertTo-AccessReviewHtmlText ($config.Scope | ConvertTo-Json -Depth 10))</pre></details>")
		}
		else {
			$null = $body.AppendLine('<p class="notice">Detailed configuration unavailable. No durations, reviewer routing or live stage dates have been inferred from names, old defaults or current settings.</p>')
			if ($result.StageNames) { $null = $body.AppendLine("<p>Recorded stage names only: $(ConvertTo-AccessReviewHtmlText $result.StageNames)</p>") }
		}
		$null = $body.AppendLine('<details><summary>All recorded discovery paths</summary><ul>')
		foreach ($path in @($result.GroupPaths)) {
			$null = $body.AppendLine("<li>$(ConvertTo-AccessReviewHtmlText (@($path.GroupNames) -join ' -> '))<br><code>$(ConvertTo-AccessReviewHtmlText (@($path.GroupIds) -join ' -> '))</code> | Application: $(ConvertTo-AccessReviewHtmlText $path.ServicePrincipalId)</li>")
		}
		if (-not @($result.GroupPaths).Count) { $null = $body.AppendLine('<li>Unavailable or not applicable.</li>') }
		$null = $body.AppendLine('</ul></details><details><summary>Recorded IDs and diagnostics</summary><dl>')
		foreach ($field in @('TargetType', 'ReviewId', 'GroupId', 'ServicePrincipalId', 'PrincipalId', 'PrincipalType', 'AssignmentId', 'Role', 'MatchedReviewName', 'Reason')) {
			$null = $body.AppendLine("<dt>$field</dt><dd>$(ConvertTo-AccessReviewHtmlText $result.$field)</dd>")
		}
		$null = $body.AppendLine('</dl></details></article>')
	}
	$null = $body.AppendLine('</section><section id="inventory" class="panel"><h2>Review inventory</h2><div class="filters"><label>Search<input id="search" type="search" placeholder="Group, reviewer ID, role, error..."></label><label>Outcome<select id="status"><option value="">All outcomes</option>')
	foreach ($status in @($outcomes.Keys | Sort-Object)) {
		$text = ConvertTo-AccessReviewHtmlText $status
		$null = $body.AppendLine("<option value=`"$text`">$text</option>")
	}
	$null = $body.AppendLine('</select></label><label>Target type<select id="type"><option value="">All types</option>')
	foreach ($type in @($Results.TargetType | Where-Object { $_ } | Sort-Object -Unique)) {
		$text = ConvertTo-AccessReviewHtmlText $type
		$null = $body.AppendLine("<option value=`"$text`">$text</option>")
	}
	$null = $body.AppendLine('</select></label></div><p id="matches" role="status"></p><div class="table-wrap"><table><thead><tr><th>Review / target</th><th>Population</th><th>Stages</th><th>Outcome</th><th>Notes</th></tr></thead><tbody id="rows">')
	for ($i = 0; $i -lt $Results.Count; $i++) {
		$result = $Results[$i]
		$name = if ($result.ReviewName) { $result.ReviewName } elseif ($result.Target) { $result.Target } else { 'Unnamed recorded item' }
		$outcome = ConvertTo-AccessReviewHtmlText (Get-AccessReviewReportOutcome $result)
		$type = ConvertTo-AccessReviewHtmlText $result.TargetType
		$population = if ($null -ne $result.MemberCount) {
			if ($result.GroupCountBasis -eq 'DirectUsers') { "$($result.MemberCount) direct users" }
			else { "$($result.MemberCount) recorded members (basis unknown)" }
		}
		elseif ($null -ne $result.SubjectCount) { "$($result.SubjectCount) directly assigned users" }
		else { 'Unavailable / not applicable' }
		$notes = @(
			if ($result.IsNested) { 'Nested group' }
			if ($result.IsSyncedGroup -eq $true) { 'Synced / report only' }
			if ($result.MatchedReviewName) { "Matched: $($result.MatchedReviewName). Existing configuration not fetched." }
			$result.Reason
		) -join ' '
		$searchText = ConvertTo-AccessReviewHtmlText (($result | ConvertTo-Json -Depth 10 -Compress))
		$groupLink = if ($result.GroupId) { New-AccessReviewGroupLink -GroupId $result.GroupId } else { '' }
		$null = $body.AppendLine("<tr data-status=`"$outcome`" data-type=`"$type`" data-search=`"$searchText`"><td><button class=`"link view`" data-target=`"review-$i`">$(ConvertTo-AccessReviewHtmlText $name)</button><small>$type</small>$groupLink</td><td>$(ConvertTo-AccessReviewHtmlText $population)</td><td>$(ConvertTo-AccessReviewHtmlText $result.StageNames)</td><td><span class=`"badge`">$outcome</span></td><td>$(ConvertTo-AccessReviewHtmlText $notes)</td></tr>")
	}
	$null = $body.AppendLine('</tbody></table></div><div id="pagination"><button id="previous" type="button">Previous</button><span id="page"></span><button id="next" type="button">Next</button></div><noscript><p>JavaScript is required for filtering, pagination, theme switching and selecting details. All inventory rows remain visible without JavaScript.</p></noscript></section><footer>Offline report. Existing reviews were not modified by report generation. Counts describe discovery-time populations, not guaranteed current membership.</footer></main>')
	$template = @'
<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'">
<title>Access review campaign</title>
<style>
:root{color-scheme:light;--bg:#f8f3f3;--panel:#fff;--ink:#302327;--muted:#715e65;--line:#e7d8dc;--accent:#b51f3b;--soft:#fbe9ed;--warning:#775015;--warn-bg:#fff3dc}
:root[data-theme=dark]{color-scheme:dark;--bg:#171315;--panel:#241d20;--ink:#f7edef;--muted:#c9b5bd;--line:#4c3941;--accent:#ff8da3;--soft:#472832;--warning:#ffda90;--warn-bg:#392e1c}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 system-ui,sans-serif}
main{max-width:1250px;margin:auto;padding:28px 24px 60px}header{display:flex;justify-content:space-between;align-items:start;gap:14px}
h1{font-size:32px;letter-spacing:-.7px;margin:2px 0 12px}h2{font-size:23px;margin-top:0}h3{font-size:21px}h4{font-size:16px}
.eyebrow{font-size:11px;letter-spacing:1.5px;text-transform:uppercase;font-weight:750;color:var(--accent);margin-bottom:0}
.theme-control{display:flex;flex-direction:column;align-items:flex-end;max-width:240px}#theme-notice{font-size:11px}
button,input,select{font:inherit;color:var(--ink);background:var(--panel);border:1px solid var(--line);border-radius:7px;padding:8px 12px}
button{cursor:pointer}button:disabled{opacity:.5;cursor:default}#theme{border-color:var(--accent);white-space:nowrap}#theme[aria-checked=true]{background:var(--soft)}
a,.link{color:var(--accent)}.link{border:0;background:transparent;padding:0;text-align:left;text-decoration:underline;overflow-wrap:anywhere}
button:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible{outline:3px solid var(--accent);outline-offset:3px}
.notice{border-left:4px solid var(--accent);background:var(--soft);border-radius:5px;padding:12px 16px}
.warning{color:var(--warning);background:var(--warn-bg);border-color:var(--warning)}
.muted,small,footer{color:var(--muted)}small{display:block}.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:22px 0}
.metric,.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px}.metric{padding:12px 16px}.metric strong{display:block;font-size:26px;color:var(--accent)}.metric span{font-size:12px}
.panel{padding:24px;margin-top:22px;scroll-margin-top:15px}nav{display:flex;gap:22px;flex-wrap:wrap}
.entra-link{display:block;font-size:12px;margin-top:6px}.topology-controls{display:flex;align-items:end;justify-content:space-between;gap:14px;flex-wrap:wrap}.topology-controls label{font-size:12px;max-width:100%}.topology-controls select{display:block;max-width:100%;margin-top:5px}.topology-controls button{padding:6px 10px}
#topology-viewport{max-height:700px;overflow:auto;border:1px solid var(--line);border-radius:8px;background:var(--bg)}#topology-space{position:relative}#topology-canvas{position:relative;transform-origin:top left}#topology-lines{position:absolute;inset:0;pointer-events:none;overflow:visible}.topology-edge{stroke:var(--accent);stroke-width:2;fill:none}#topology-arrow path{fill:var(--accent)}
.topology-node{width:260px;max-width:100%;min-height:180px;padding:14px;margin:12px;background:var(--panel);border:1px solid var(--line);border-top:4px solid var(--accent);border-radius:9px;display:flex;flex-direction:column;gap:6px;overflow-wrap:anywhere}.topology-node.positioned{position:absolute;margin:0}.topology-node .eyebrow{font-size:10px;letter-spacing:.5px}.topology-name{font-weight:700;font-size:15px}.topology-stages{font-size:12px}.topology-name,.topology-stages{display:-webkit-box;-webkit-box-orient:vertical;overflow:hidden;-webkit-line-clamp:3}.topology-node .entra-link{margin-top:auto}
.relationships{display:grid;grid-template-columns:1fr 1fr 1fr;gap:15px;padding:12px;border:1px solid var(--line);border-radius:7px}.relationships>div{min-width:0;overflow-wrap:anywhere}.selected-group{display:flex;flex-direction:column;justify-content:center;gap:8px;padding:15px;background:var(--soft);border:1px solid var(--accent);border-radius:7px}
.timeline{display:flex;gap:3px;margin:18px 0}.stage{flex-basis:0;min-width:0;background:var(--soft);border-top:4px solid var(--accent);padding:12px;border-radius:6px;overflow-wrap:anywhere}
.stage strong,.stage span{display:block}.stage small{font-size:11px}details{margin:14px 0}summary{cursor:pointer;color:var(--accent);font-weight:600}
dl{display:grid;grid-template-columns:minmax(110px,1fr) 3fr;gap:5px 15px}dt{font-weight:600}dd{margin:0;overflow-wrap:anywhere}
pre,code{white-space:pre-wrap;overflow-wrap:anywhere;font-size:12px}.filters{display:flex;gap:14px;flex-wrap:wrap}.filters label{flex:1;min-width:170px;font-size:12px;font-weight:600}.filters input,.filters select{display:block;width:100%;margin-top:5px}
.table-wrap{overflow:auto}table{width:100%;border-collapse:collapse;font-size:13px}th,td{text-align:left;padding:13px 10px;border-bottom:1px solid var(--line);vertical-align:top}th{font-size:11px;text-transform:uppercase;color:var(--muted)}td{overflow-wrap:anywhere;max-width:350px}.badge{display:inline-block;background:var(--soft);color:var(--accent);padding:3px 9px;border-radius:18px;font-size:12px}
#pagination{display:flex;justify-content:center;gap:18px;align-items:center;margin-top:18px}footer{margin-top:24px;font-size:12px}[hidden]{display:none!important}
@media(prefers-color-scheme:dark){:root:not([data-theme]){color-scheme:dark;--bg:#171315;--panel:#241d20;--ink:#f7edef;--muted:#c9b5bd;--line:#4c3941;--accent:#ff8da3;--soft:#472832;--warning:#ffda90;--warn-bg:#392e1c}}
@media(max-width:650px){main{padding:18px 12px}h1{font-size:25px}.panel{padding:15px}.metrics{grid-template-columns:1fr 1fr}.timeline{flex-direction:column}.relationships{grid-template-columns:1fr}.theme-control{max-width:130px}dl{grid-template-columns:1fr}#pagination{gap:10px}}
@media print{:root,:root[data-theme=dark]{color-scheme:light;--bg:#fff;--panel:#fff;--ink:#222;--muted:#555;--line:#ccc;--accent:#931b30;--soft:#fff;--warning:#604000;--warn-bg:#fff}body{background:white}main{max-width:none;padding:0}.theme-control,.filters,#pagination,nav,.topology-controls{display:none}.panel{border:0}.table-wrap{overflow:visible}tr[hidden]{display:table-row!important}.review-detail[hidden]{display:none!important}#topology-viewport{max-height:none;overflow:visible}#topology-space,#topology-canvas{width:auto!important;height:auto!important;transform:none!important}#topology-lines{display:none}.topology-node.positioned{position:static;display:inline-flex;vertical-align:top;margin:6px;break-inside:avoid}h2,h3{break-after:avoid}}
</style></head><body>
<!--REPORT-->
<script>
(() => {
  const root = document.documentElement;
  const toggle = document.getElementById('theme');
  const notice = document.getElementById('theme-notice');
  const system = matchMedia('(prefers-color-scheme: dark)');
  const key = 'access-review-report-theme';
  let choice = null;
  const storageProblem = error => {
    if (error.name !== 'SecurityError' && error.name !== 'QuotaExceededError') throw error;
    notice.textContent = 'Theme preference cannot be saved in this browser context.';
  };
  try { const saved = localStorage.getItem(key); if (saved === 'light' || saved === 'dark') choice = saved; }
  catch (error) { storageProblem(error); }
  function applyTheme() {
    const dark = choice ? choice === 'dark' : system.matches;
    root.dataset.theme = dark ? 'dark' : 'light';
    toggle.setAttribute('aria-checked', String(dark));
  }
  applyTheme();
  system.addEventListener('change', () => { if (!choice) applyTheme(); });
  toggle.addEventListener('click', () => {
    choice = root.dataset.theme === 'dark' ? 'light' : 'dark';
    applyTheme();
    try { localStorage.setItem(key, choice); } catch (error) { storageProblem(error); }
  });
  function initializeTopology() {
    const nodes = [...document.querySelectorAll('.topology-node')];
    const edges = [...document.querySelectorAll('.topology-edge')];
    const viewport = document.getElementById('topology-viewport');
    const canvas = document.getElementById('topology-canvas');
    const space = document.getElementById('topology-space');
    const lines = document.getElementById('topology-lines');
    const selector = document.getElementById('topology-component');
    const counter = document.getElementById('topology-count');
    if (!nodes.length) {
      counter.textContent = 'No group IDs recorded; group topology is unavailable for this report.';
      viewport.hidden = true;
      for (const control of document.querySelectorAll('.topology-controls button, .topology-controls select')) control.disabled = true;
      return;
    }
    const byId = new Map(nodes.map(node => [node.id, node]));
    const incoming = new Map(nodes.map(node => [node.id, 0]));
    const outgoing = new Map(nodes.map(node => [node.id, []]));
    const neighbors = new Map(nodes.map(node => [node.id, new Set()]));
    for (const edge of edges) {
      const { from, to } = edge.dataset;
      incoming.set(to, incoming.get(to) + 1);
      outgoing.get(from).push(to);
      neighbors.get(from).add(to);
      neighbors.get(to).add(from);
    }
    const components = [];
    const visited = new Set();
    for (const node of nodes) {
      if (visited.has(node.id)) continue;
      const queue = [node.id];
      visited.add(node.id);
      for (let i = 0; i < queue.length; i++) {
        for (const id of neighbors.get(queue[i])) {
          if (!visited.has(id)) { visited.add(id); queue.push(id); }
        }
      }
      components.push(new Set(queue));
    }
    components.forEach((ids, index) => {
      const first = [...ids].find(id => incoming.get(id) === 0) || ids.values().next().value;
      const option = document.createElement('option');
      option.value = String(index);
      option.textContent = `${index + 1}. ${byId.get(first).querySelector('.topology-name').textContent.trim()} (${ids.size} ${ids.size === 1 ? 'group' : 'groups'})`;
      selector.append(option);
    });
    const levels = new Map(nodes.map(node => [node.id, 0]));
    const remaining = new Map(incoming);
    const queue = nodes.filter(node => incoming.get(node.id) === 0).map(node => node.id);
    for (let i = 0; i < queue.length; i++) {
      for (const id of outgoing.get(queue[i])) {
        levels.set(id, Math.max(levels.get(id), levels.get(queue[i]) + 1));
        remaining.set(id, remaining.get(id) - 1);
        if (remaining.get(id) === 0) queue.push(id);
      }
    }
    const affected = new Set(nodes.filter(node => remaining.get(node.id) > 0).map(node => node.id));
    const lastLevel = Math.max(...levels.values()) + 1;
    for (const id of affected) levels.set(id, lastLevel);
    let width = 0, height = 0, zoom = 1, summary = '', lastWidth = 0;
    function setZoom(value) {
      const minimum = Math.min(.1, Math.max(1, viewport.clientWidth) / width);
      zoom = Math.max(minimum, Math.min(2, value));
      canvas.style.transform = `scale(${zoom})`;
      space.style.width = `${width * zoom}px`;
      space.style.height = `${height * zoom}px`;
      counter.textContent = `${summary} Zoom ${Number((zoom * 100).toFixed(1))}%.`;
      document.getElementById('zoom-out').disabled = zoom <= minimum;
      document.getElementById('zoom-in').disabled = zoom >= 2;
    }
    function layout(fit) {
      const visible = selector.value === 'all' ? new Set(byId.keys()) : components[Number(selector.value)];
      const rows = new Map();
      for (const node of nodes) {
        node.hidden = !visible.has(node.id);
        node.classList.add('positioned');
        if (node.hidden) continue;
        const level = levels.get(node.id);
        if (!rows.has(level)) rows.set(level, []);
        rows.get(level).push(node);
      }
      width = Math.max(viewport.clientWidth, ...[...rows.values()].map(row => row.length * 292 + 16));
      canvas.style.width = `${width}px`;
      const positions = new Map();
      let top = 24;
      for (const [, row] of [...rows].sort((a, b) => a[0] - b[0])) {
        const left = (width - (row.length * 292 - 32)) / 2;
        const rowHeight = Math.max(...row.map(node => node.offsetHeight));
        row.forEach((node, index) => {
          const x = left + index * 292;
          node.style.left = `${x}px`;
          node.style.top = `${top}px`;
          positions.set(node.id, { x: x + node.offsetWidth / 2, right: x + node.offsetWidth, top, bottom: top + node.offsetHeight });
        });
        top += rowHeight + 80;
      }
      height = top - 56;
      canvas.style.height = `${height}px`;
      lines.setAttribute('width', width);
      lines.setAttribute('height', height);
      let visibleEdges = 0;
      for (const edge of edges) {
        const from = positions.get(edge.dataset.from), to = positions.get(edge.dataset.to);
        const hidden = !from || !to;
        edge.toggleAttribute('hidden', hidden);
        if (hidden) { edge.removeAttribute('d'); continue; }
        visibleEdges++;
        if (edge.dataset.from === edge.dataset.to) {
          const upper = from.top + 24, lower = from.bottom - 24;
          edge.setAttribute('d', `M ${from.right} ${lower} C ${from.right + 28} ${lower}, ${from.right + 28} ${upper}, ${from.right} ${upper}`);
          continue;
        }
        const bend = (from.bottom + to.top) / 2;
        edge.setAttribute('d', to.top > from.bottom
          ? `M ${from.x} ${from.bottom} C ${from.x} ${bend}, ${to.x} ${bend}, ${to.x} ${to.top}`
          : `M ${from.x} ${from.bottom} C ${from.x + 100} ${from.bottom + 50}, ${to.x + 100} ${to.top - 50}, ${to.x} ${to.top}`);
      }
      const warning = document.getElementById('topology-warning');
      warning.hidden = ![...visible].some(id => affected.has(id));
      warning.textContent = 'Recorded connections contain a cycle. Affected nodes are placed separately; connections are retained. This is not a valid hierarchical ordering.';
      summary = `${visible.size} of ${nodes.length} recorded groups; ${visibleEdges} of ${edges.length} connections. ${components.length} connected hierarchies.`;
      lastWidth = viewport.clientWidth;
      setZoom(fit ? Math.max(.4, Math.min(1, viewport.clientWidth / width)) : zoom);
    }
    // Keep large disconnected campaigns navigable without dropping any groups.
    if (nodes.length > 50 && components.length > 1) selector.value = '0';
    selector.addEventListener('change', () => { layout(true); viewport.scrollTo(0, 0); });
    document.getElementById('zoom-in').addEventListener('click', () => setZoom(zoom + .1));
    document.getElementById('zoom-out').addEventListener('click', () => setZoom(zoom - .1));
    document.getElementById('zoom-fit').addEventListener('click', () => setZoom(Math.min(1, viewport.clientWidth / width)));
    document.getElementById('zoom-reset').addEventListener('click', () => setZoom(1));
    new ResizeObserver(() => {
      if (viewport.clientWidth !== lastWidth) layout(false);
    }).observe(viewport);
    layout(true);
  }
  initializeTopology();
  const search = document.getElementById('search');
  const status = document.getElementById('status');
  const type = document.getElementById('type');
  const rows = [...document.querySelectorAll('#rows tr')];
  const pageSize = 50;
  let current = 0;
  let filtered = rows;
  function render() {
    const pages = Math.max(1, Math.ceil(filtered.length / pageSize));
    current = Math.min(current, pages - 1);
    for (const row of rows) row.hidden = true;
    for (const row of filtered.slice(current * pageSize, (current + 1) * pageSize)) row.hidden = false;
    document.getElementById('matches').textContent = `${filtered.length} of ${rows.length} recorded targets`;
    document.getElementById('page').textContent = `Page ${current + 1} of ${pages}`;
    document.getElementById('previous').disabled = current === 0;
    document.getElementById('next').disabled = current === pages - 1;
  }
  function filter() {
    const query = search.value.trim().toLowerCase();
    filtered = rows.filter(row => (!status.value || row.dataset.status === status.value) && (!type.value || row.dataset.type === type.value) && row.dataset.search.toLowerCase().includes(query));
    current = 0;
    render();
  }
  search.addEventListener('input', filter);
  status.addEventListener('change', filter);
  type.addEventListener('change', filter);
  document.getElementById('previous').addEventListener('click', () => { current--; render(); });
  document.getElementById('next').addEventListener('click', () => { current++; render(); });
  document.addEventListener('click', event => {
    const button = event.target.closest('button.view');
    if (!button) return;
    for (const detail of document.querySelectorAll('.review-detail')) detail.hidden = detail.id !== button.dataset.target;
    document.getElementById('selection-hint').hidden = true;
    const target = document.getElementById(button.dataset.target);
    target.querySelector('h3').focus();
    target.scrollIntoView({ block: 'start' });
  });
  render();
})();
</script></body></html>
'@
	$template.Replace('<!--REPORT-->', $body.ToString())
}

function Write-AccessReviewReport {
	param(
		[Parameter(Mandatory)][string]$Path,
		[AllowEmptyCollection()][object[]]$Results = @(),
		[Parameter(Mandatory)][hashtable]$Context
	)

	$fullPath = Resolve-AccessReviewReportPath -Path $Path
	$html = New-AccessReviewHtml -Results $Results -Context $Context
	$directory = [System.IO.Path]::GetDirectoryName($fullPath)
	$temporaryPath = Join-Path $directory ('.access-review-' + [guid]::NewGuid().ToString() + '.tmp')
	try {
		$null = [System.IO.Directory]::CreateDirectory($directory)
		[System.IO.File]::WriteAllText($temporaryPath, $html, [System.Text.UTF8Encoding]::new($false))
		[System.IO.File]::Move($temporaryPath, $fullPath)
	}
	finally {
		if ([System.IO.File]::Exists($temporaryPath)) { [System.IO.File]::Delete($temporaryPath) }
	}
	Write-Host "HTML report: $fullPath"
}
