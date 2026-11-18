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
	if ($Result.DuplicateMatchType -eq 'Resource') { 'Existing resource match' }
	elseif ($Result.DuplicateMatchType -eq 'UncertainResourceScope') { 'Manual review required' }
	elseif ($Result.MatchedReviewName) { 'Existing name match' }
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

function New-AccessReviewApplicationLink {
	param([AllowNull()][string]$ServicePrincipalId)
	$id = [guid]::Empty
	if ([guid]::TryParse($ServicePrincipalId, [ref]$id) -and $id -ne [guid]::Empty) {
		$url = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$id"
		"<a class=`"entra-link`" href=`"$url`" target=`"_blank`" rel=`"noopener noreferrer`">Open application in Entra</a>"
	}
	else { '<small class="muted">Entra link unavailable: valid service-principal Object ID not recorded.</small>' }
}

function Get-AccessReviewReportStyle {
	@'
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
.status-dashboard-layout{display:flex;align-items:center;gap:32px;flex-wrap:wrap}.status-pie{width:220px;aspect-ratio:1;border:1px solid var(--line);border-radius:50%;flex:0 0 220px}.status-legend{list-style:none;padding:0;margin:0;min-width:220px;flex:1}.status-legend li{display:grid;grid-template-columns:16px minmax(0,1fr) auto auto;gap:12px;align-items:center;padding:8px 0;border-bottom:1px solid var(--line)}.status-label{overflow-wrap:anywhere}.status-swatch{width:12px;height:12px;border:1px solid var(--line);border-radius:3px}
.panel{padding:24px;margin-top:22px;scroll-margin-top:15px}nav{display:flex;gap:22px;flex-wrap:wrap}
.entra-link{display:block;font-size:12px;margin-top:6px}.topology-controls{display:flex;align-items:end;justify-content:space-between;gap:14px;flex-wrap:wrap}.topology-controls label{font-size:12px;max-width:100%}.topology-controls select{display:block;max-width:100%;margin-top:5px}.topology-controls button{padding:6px 10px}
#topology-viewport{max-height:700px;overflow:auto;border:1px solid var(--line);border-radius:8px;background:var(--bg)}#topology-space{position:relative}#topology-canvas{position:relative;transform-origin:top left}#topology-lines{position:absolute;inset:0;pointer-events:none;overflow:visible}.topology-edge{stroke:var(--accent);stroke-width:2;fill:none}#topology-arrow path{fill:var(--accent)}
.topology-node{width:260px;max-width:100%;min-height:180px;padding:14px;margin:12px;background:var(--panel);border:1px solid var(--line);border-top:4px solid var(--accent);border-radius:9px;display:flex;flex-direction:column;gap:6px;overflow-wrap:anywhere}.topology-node.positioned{position:absolute;margin:0}.topology-node .eyebrow{font-size:10px;letter-spacing:.5px}.topology-name{font-weight:700;font-size:15px}.topology-stages{font-size:12px}.topology-name,.topology-stages{display:-webkit-box;-webkit-box-orient:vertical;overflow:hidden;-webkit-line-clamp:3}.topology-node .entra-link{margin-top:auto}
.application-node{background:var(--soft)}.application-edge{stroke-dasharray:6 4}
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
'@
}

function Get-AccessReviewReportThemeScript {
	@'
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
'@
}

function New-AccessReviewHtml {
	param(
		[AllowEmptyCollection()][object[]]$Results = @(),
		[Parameter(Mandatory)][hashtable]$Context
	)

	$Results = @(
		$Results | Where-Object { $_.TargetType -ne 'ServicePrincipalOwner' -or -not $_.ServicePrincipalId }
		foreach ($owners in @($Results | Where-Object { $_.TargetType -eq 'ServicePrincipalOwner' -and $_.ServicePrincipalId } |
			Group-Object -Property ServicePrincipalId, Status)) {
			$first = $owners.Group[0]
			[pscustomobject]@{
				TargetType = 'ServicePrincipalOwner'; Target = $first.Target
				ServicePrincipalId = $first.ServicePrincipalId
				ServicePrincipalName = $first.ServicePrincipalName
				Status = $first.Status; Reason = $first.Reason; Role = $first.Role
				OwnerItems = @($owners.Group)
			}
		}
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
	foreach ($status in @('Created', 'Would be created', 'Existing name match', 'Existing resource match', 'Failed', 'Skipped', 'Not created', 'Manual review required', 'Unknown outcome')) {
		$count = if ($outcomes.ContainsKey($status)) { $outcomes[$status] } else { 0 }
		$null = $body.AppendLine("<div class=`"metric`"><strong>$count</strong><span>$(ConvertTo-AccessReviewHtmlText $status)</span></div>")
	}
	$null = $body.AppendLine('</section>')
	if (@($Results | Where-Object OwnerItems).Count) {
		$null = $body.AppendLine('<p class="muted">Ownership follow-up is grouped by application and outcome, not counted once per owner. Expand an entry to see its recorded owners. Raw results and logs retain individual ownership records.</p>')
	}
	$groups = @($Results | Where-Object { $_.GroupId -and $_.TargetType -notin @('ApplicationDirectUsers', 'ServicePrincipalOwner', 'ApplicationNonUserAssignment') })
	$knownGroups = @($groups | Where-Object { $_.GroupCountBasis -eq 'DirectUsers' -and $null -ne $_.MemberCount })
	$unknownGroups = @($groups | Where-Object { $_.GroupCountBasis -ne 'DirectUsers' -and $null -ne $_.MemberCount })
	$groupTotal = [long](($knownGroups | Measure-Object MemberCount -Sum).Sum)
	$appTotal = [long](($Results | Where-Object TargetType -eq 'ApplicationDirectUsers' | Measure-Object SubjectCount -Sum).Sum)
	$null = $body.AppendLine("<p class=`"muted`">All recorded targets: <strong>$groupTotal</strong> verified direct-user counts summed across group targets; <strong>$appTotal</strong> direct-application user counts, separately. These are per-review populations, not unique people. <strong>$($unknownGroups.Count)</strong> group count(s) have an unknown historical basis and are excluded from the direct-user total.</p>")
	$null = $body.AppendLine('<nav><a href="#topology">Resource topology</a><a href="#configuration">Stages and details</a><a href="#inventory">Search inventory</a></nav>')
	$applications = @{}
	$applicationSources = @(
		$Context.Applications | Where-Object { $_ }
		for ($i = 0; $i -lt $Results.Count; $i++) {
			$result = $Results[$i]
			if ($result.ServicePrincipalId) {
				[pscustomobject]@{
					ServicePrincipalId = $result.ServicePrincipalId
					ServicePrincipalName = if ($result.ServicePrincipalName) { $result.ServicePrincipalName }
						elseif ($result.TargetType -ne 'GroupMembership') { $result.Target }
					DirectUserIndex = if ($result.TargetType -eq 'ApplicationDirectUsers') { $i } else { $null }
				}
			}
			foreach ($path in @($result.GroupPaths | Where-Object ServicePrincipalId)) {
				[pscustomobject]@{
					ServicePrincipalId = $path.ServicePrincipalId; ServicePrincipalName = $path.ServicePrincipalName
					GroupIds = @($path.GroupIds | Select-Object -First 1)
				}
			}
			foreach ($sourceId in @($result.SourceServicePrincipalIds | Where-Object { $_ })) {
				[pscustomobject]@{ ServicePrincipalId = $sourceId }
			}
		}
	)
	foreach ($source in $applicationSources) {
		$id = [string]$source.ServicePrincipalId
		if (-not $id) { continue }
		if (-not $applications.ContainsKey($id)) {
			$applications[$id] = @{
				Name = $id; DirectUserIndex = $null; DirectUserCount = $null
				GroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
			}
		}
		$app = $applications[$id]
		if ($source.ServicePrincipalName) { $app.Name = $source.ServicePrincipalName }
		if ($null -ne $source.DirectUserIndex) { $app.DirectUserIndex = $source.DirectUserIndex }
		if ($null -ne $source.DirectUserCount) { $app.DirectUserCount = $source.DirectUserCount }
		foreach ($groupId in @($source.GroupIds | Where-Object { $_ })) { $null = $app.GroupIds.Add([string]$groupId) }
	}
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
	$assignedGroupIds = @($applications.Values | ForEach-Object { $_.GroupIds })
	$topologyIds = @(@($groupIndex.Keys) + @($parents.Keys) + @($children.Keys) + $assignedGroupIds | Sort-Object -Unique)
	$applicationIds = @($applications.Keys | Sort-Object)
	$applicationIndex = @{}
	for ($i = 0; $i -lt $applicationIds.Count; $i++) { $applicationIndex[$applicationIds[$i]] = $i }
	$topologyIndex = @{}
	for ($i = 0; $i -lt $topologyIds.Count; $i++) { $topologyIndex[$topologyIds[$i]] = $i }
	$null = $body.AppendLine('<section id="topology" class="panel"><h2>Application and group topology</h2><p>Dashed arrows connect enterprise applications to their directly assigned groups. Solid arrows show recorded parent-to-child group membership, not guaranteed effective application access. Shared resources appear once. Application cards show direct-user review outcomes separately from group reviews.</p><div class="topology-controls"><label>Connected hierarchy<select id="topology-component" aria-controls="topology-viewport"><option value="all">All recorded resources</option></select></label><div><button id="zoom-out" type="button" aria-label="Zoom topology out">-</button> <button id="zoom-in" type="button" aria-label="Zoom topology in">+</button> <button id="zoom-fit" type="button">Fit width</button> <button id="zoom-reset" type="button">100%</button></div></div><p id="topology-count" class="muted" role="status"></p><p id="topology-warning" class="notice warning" role="status" hidden></p>')
	if ($incompleteTopology) { $null = $body.AppendLine('<p class="notice warning">Some recorded path IDs are missing; those incomplete connections cannot be drawn. Other recorded nodes and connections remain visible.</p>') }
	$null = $body.AppendLine('<div id="topology-viewport" tabindex="0" aria-label="Scrollable application and group topology"><div id="topology-space"><div id="topology-canvas"><svg id="topology-lines" aria-hidden="true"><defs><marker id="topology-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z"/></marker></defs>')
	foreach ($appId in $applicationIds) {
		foreach ($groupId in @($applications[$appId].GroupIds | Sort-Object)) {
			$null = $body.AppendLine("<path class=`"topology-edge application-edge`" data-from=`"app-$($applicationIndex[$appId])`" data-to=`"node-$($topologyIndex[$groupId])`" marker-end=`"url(#topology-arrow)`"><title>Direct application-role group assignment</title></path>")
		}
	}
	foreach ($childId in @($parents.Keys | Sort-Object)) {
		foreach ($parentId in @($parents[$childId] | Sort-Object)) {
			$null = $body.AppendLine("<path class=`"topology-edge`" data-from=`"node-$($topologyIndex[$parentId])`" data-to=`"node-$($topologyIndex[$childId])`" marker-end=`"url(#topology-arrow)`"/>")
		}
	}
	$null = $body.AppendLine('</svg>')
	foreach ($appId in $applicationIds) {
		$app = $applications[$appId]
		$null = $body.AppendLine("<article id=`"app-$($applicationIndex[$appId])`" class=`"topology-node application-node`"><small class=`"eyebrow`">Enterprise application / service principal</small>")
		if ($null -ne $app.DirectUserIndex) {
			$result = $Results[$app.DirectUserIndex]
			$population = if ($null -ne $result.SubjectCount) { "$($result.SubjectCount) directly assigned users" } else { 'Direct-user population unavailable' }
			$null = $body.AppendLine("<button class=`"link view topology-name`" data-target=`"review-$($app.DirectUserIndex)`">$(ConvertTo-AccessReviewHtmlText $app.Name)</button><span>$(ConvertTo-AccessReviewHtmlText $population)</span><small>Direct-user review: $(ConvertTo-AccessReviewHtmlText (Get-AccessReviewReportOutcome $result))</small><small class=`"topology-stages`">$(ConvertTo-AccessReviewHtmlText $result.StageNames)</small>")
		}
		else {
			$null = $body.AppendLine("<strong class=`"topology-name`">$(ConvertTo-AccessReviewHtmlText $app.Name)</strong><small>$(if ($null -ne $app.DirectUserCount) { "$($app.DirectUserCount) directly assigned users discovered. " })No direct-user review result recorded.</small>")
		}
		$null = $body.AppendLine("$(New-AccessReviewApplicationLink $appId)</article>")
	}
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
	$null = $body.AppendLine('</div></div></div><noscript><p>JavaScript is required to lay out the connected topology. Resource cards remain listed without connectors.</p></noscript><p class="muted">Connections reflect recorded discovery, not live membership. Resources without recorded connections are not assumed to be unrelated. Entra links require signing in to the appropriate directory with permission to view the resource.</p></section><section id="configuration" class="panel"><h2>Hierarchy and stage configuration</h2><p id="selection-hint" class="muted">Choose an application or group in the topology, or a review in the inventory, to see its relationships, paths and available configuration.</p>')
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
		if ($result.OwnerItems) {
			$null = $body.AppendLine("<p class=`"notice warning`">Ownership requires manual review; these owners are not subjects certified by the application user-access review. No ownership changes are applied.</p><details><summary>Recorded owners ($($result.OwnerItems.Count))</summary><div class=`"table-wrap`"><table><thead><tr><th>Owner</th><th>Type</th><th>Object ID</th><th>Notes</th></tr></thead><tbody>")
			foreach ($owner in $result.OwnerItems) {
				$name = if ($owner.PrincipalDisplayName) { $owner.PrincipalDisplayName } else { 'Name not recorded' }
				$null = $body.AppendLine("<tr><td>$(ConvertTo-AccessReviewHtmlText $name)</td><td>$(ConvertTo-AccessReviewHtmlText $owner.PrincipalType)</td><td>$(ConvertTo-AccessReviewHtmlText $owner.PrincipalId)</td><td>$(ConvertTo-AccessReviewHtmlText $owner.Reason)</td></tr>")
			}
			$null = $body.AppendLine('</tbody></table></div></details>')
		}
		elseif ($config -and @($config.Stages).Count) {
			$label = if ($result.Status -eq 'Created') { 'Submitted configuration (not fetched live)' } else { 'Computed configuration (not evidence of creation)' }
			$null = $body.AppendLine("<h4>$label</h4><p>Frequency: $(ConvertTo-AccessReviewHtmlText $config.Frequency). First planned instance: $(ConvertTo-AccessReviewHtmlText $config.StartDate), $([int]$config.InstanceDurationInDays) days.</p>")
			if ($config.DescriptionForReviewers) {
				$null = $body.AppendLine("<p><strong>Reviewer description:</strong> $(ConvertTo-AccessReviewHtmlText $config.DescriptionForReviewers)</p>")
			}
			if ($config.DescriptionForAdmins) {
				$null = $body.AppendLine("<p><strong>Administrator description:</strong> $(ConvertTo-AccessReviewHtmlText $config.DescriptionForAdmins)</p>")
			}
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
			$applicationLabel = if ($path.ServicePrincipalName) { "$($path.ServicePrincipalName) ($($path.ServicePrincipalId))" } else { $path.ServicePrincipalId }
			$null = $body.AppendLine("<li>$(ConvertTo-AccessReviewHtmlText (@($path.GroupNames) -join ' -> '))<br><code>$(ConvertTo-AccessReviewHtmlText (@($path.GroupIds) -join ' -> '))</code> | Application: $(ConvertTo-AccessReviewHtmlText $applicationLabel)</li>")
		}
		if (-not @($result.GroupPaths).Count) { $null = $body.AppendLine('<li>Unavailable or not applicable.</li>') }
		$null = $body.AppendLine('</ul></details><details><summary>Recorded IDs and diagnostics</summary><dl>')
		foreach ($field in @('TargetType', 'ReviewId', 'GroupId', 'ServicePrincipalId', 'PrincipalId', 'PrincipalType', 'AssignmentId', 'Role', 'DuplicateMatchType', 'MatchedReviewName', 'MatchedReviewId', 'MatchedReviewScope', 'Reason')) {
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
		$population = if ($result.OwnerItems) { "$($result.OwnerItems.Count) recorded owner$(if ($result.OwnerItems.Count -ne 1) { 's' })" }
		elseif ($null -ne $result.MemberCount) {
			if ($result.GroupCountBasis -eq 'DirectUsers') { "$($result.MemberCount) direct users" }
			else { "$($result.MemberCount) recorded members (basis unknown)" }
		}
		elseif ($null -ne $result.SubjectCount) { "$($result.SubjectCount) directly assigned users" }
		else { 'Unavailable / not applicable' }
		$notes = @(
			if ($result.IsNested) { 'Nested group' }
			if ($result.IsSyncedGroup -eq $true) { 'Synced / report only' }
			if ($result.MatchedReviewName) {
				if ($result.DuplicateMatchType -in @('Resource', 'UncertainResourceScope')) {
					"Matched: $($result.MatchedReviewName) ($($result.MatchedReviewId)). Existing stage/reviewer policy not compared or changed."
				}
				else { "Matched: $($result.MatchedReviewName). Existing configuration not fetched." }
			}
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
<!--STYLE-->
</style></head><body>
<!--REPORT-->
<script>
(() => {
  /*THEME*/
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
      counter.textContent = 'No application or group IDs recorded; resource topology is unavailable for this report.';
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
      option.textContent = `${index + 1}. ${byId.get(first).querySelector('.topology-name').textContent.trim()} (${ids.size} ${ids.size === 1 ? 'resource' : 'resources'})`;
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
      summary = `${visible.size} of ${nodes.length} recorded resources; ${visibleEdges} of ${edges.length} connections. ${components.length} connected hierarchies.`;
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
	$template.Replace('<!--STYLE-->', (Get-AccessReviewReportStyle)).Replace('/*THEME*/', (Get-AccessReviewReportThemeScript)).Replace('<!--REPORT-->', $body.ToString())
}

function Write-AccessReviewHtmlFile {
	param(
		[Parameter(Mandatory)][string]$Path,
		[Parameter(Mandatory)][string]$Html
	)

	$fullPath = Resolve-AccessReviewReportPath -Path $Path
	$directory = [System.IO.Path]::GetDirectoryName($fullPath)
	$temporaryPath = Join-Path $directory ('.access-review-' + [guid]::NewGuid().ToString() + '.tmp')
	try {
		$null = [System.IO.Directory]::CreateDirectory($directory)
		[System.IO.File]::WriteAllText($temporaryPath, $Html, [System.Text.UTF8Encoding]::new($false))
		[System.IO.File]::Move($temporaryPath, $fullPath)
	}
	finally {
		if ([System.IO.File]::Exists($temporaryPath)) { [System.IO.File]::Delete($temporaryPath) }
	}
	Write-Host "HTML report: $fullPath"
}

function Write-AccessReviewReport {
	param(
		[Parameter(Mandatory)][string]$Path,
		[AllowEmptyCollection()][object[]]$Results = @(),
		[Parameter(Mandatory)][hashtable]$Context
	)
	$fullPath = Resolve-AccessReviewReportPath -Path $Path
	Write-AccessReviewHtmlFile -Path $fullPath -Html (New-AccessReviewHtml -Results $Results -Context $Context)
}

function ConvertTo-AccessReviewStatusUsersHtml {
	param([object[]]$Users, [switch]$ShowDepartment)
	$lines = @(
		foreach ($user in @($Users | Where-Object { $_ })) {
			$department = if ($user.Error) { 'Unavailable (profile lookup failed)' }
				elseif ([string]::IsNullOrWhiteSpace($user.Department)) { 'Not set / not returned' }
				else { $user.Department }
			"<li><strong>$(ConvertTo-AccessReviewHtmlText $user.DisplayName)</strong> $(ConvertTo-AccessReviewHtmlText $user.UserPrincipalName)$(if ($ShowDepartment) { '<small class="reviewer-department">Department (current directory): ' + (ConvertTo-AccessReviewHtmlText $department) + '</small>' })<small>$(ConvertTo-AccessReviewHtmlText $user.Id) | $(ConvertTo-AccessReviewHtmlText $user.ProfileSource)</small>$(if ($user.Error) { '<span class="warning">' + (ConvertTo-AccessReviewHtmlText $user.Error) + '</span>' })</li>"
		}
	)
	if (-not $lines.Count) { return '<p class="muted">No individual user identities returned in this list.</p>' }
	'<ul>' + ($lines -join '') + '</ul>'
}

function ConvertTo-AccessReviewStatusScopesHtml {
	param([object[]]$Scopes)
	if (-not @($Scopes | Where-Object { $_ }).Count) { return '<p class="muted">None configured/returned.</p>' }
	(@(foreach ($scope in $Scopes) {
		"<p><strong>$(ConvertTo-AccessReviewHtmlText $scope.Mode)</strong><small>$(ConvertTo-AccessReviewHtmlText $scope.Note)</small></p>"
		if ($scope.Mode -in @('Self-review / default reviewer behavior', 'Line managers')) {
			'<p class="muted">Individual users are listed in the instance and runtime-stage decision tables below, when available.</p>'
		}
		else { ConvertTo-AccessReviewStatusUsersHtml $scope.Users -ShowDepartment }
		"<details><summary>Technical reviewer rule</summary><code>$(ConvertTo-AccessReviewHtmlText $scope.Query)</code>$(if ($scope.RawScope) { '<pre>' + (ConvertTo-AccessReviewHtmlText ($scope.RawScope | ConvertTo-Json -Depth 10)) + '</pre>' })</details>"
	}) -join '')
}

function ConvertTo-AccessReviewStatusOwnersHtml {
	param([object[]]$Resources)
	if (-not @($Resources | Where-Object { $_ }).Count) {
		return '<p class="notice warning">Resource IDs could not be resolved from this scope. No owner identities inferred.</p>'
	}
	(@(foreach ($resource in $Resources) {
		"<h4>$(ConvertTo-AccessReviewHtmlText $resource.DisplayName)</h4><small>$(ConvertTo-AccessReviewHtmlText $resource.Kind)/$(ConvertTo-AccessReviewHtmlText $resource.Id)</small>"
		if ($resource.Kind -eq 'groups') { New-AccessReviewGroupLink $resource.Id }
		if ($resource.MetadataError -or $resource.OwnersError) {
			"<p class=`"notice warning`">Owner/resource lookup incomplete: $(ConvertTo-AccessReviewHtmlText $resource.MetadataError) $(ConvertTo-AccessReviewHtmlText $resource.OwnersError)</p>"
		}
		ConvertTo-AccessReviewStatusUsersHtml $resource.Owners -ShowDepartment
		foreach ($other in $resource.OtherOwners) {
			"<p class=`"muted`">Non-user/unclassified owner, not listed as a human reviewer: $(ConvertTo-AccessReviewHtmlText $other.displayName) <code>$(ConvertTo-AccessReviewHtmlText $other.id)</code> $(ConvertTo-AccessReviewHtmlText $other.'@odata.type')</p>"
		}
	}) -join '')
}

function ConvertTo-AccessReviewStatusSummaryHtml {
	param([AllowNull()][object]$Summary)
	if ($null -eq $Summary) { return '<p class="notice warning">Decision counts unavailable; no zero or completion percentage inferred.</p>' }
	$percent = if ($null -ne $Summary.Completion) { "$($Summary.Completion)%" } else { 'n/a (no items)' }
	$counts = ConvertTo-AccessReviewHtmlText ($Summary.DecisionCounts | ConvertTo-Json -Compress)
	$apply = ConvertTo-AccessReviewHtmlText ($Summary.ApplyResults | ConvertTo-Json -Compress)
	"<p><strong>$($Summary.Recorded) / $($Summary.Total)</strong> recorded outcomes ($percent). Approve: $($Summary.Approve); Deny: $($Summary.Deny); Don't know: $($Summary.DontKnow); Not reviewed: $($Summary.NotReviewed); Other: $($Summary.Other).</p><details><summary>Decision and application-result breakdown</summary><pre>$counts</pre><p>Application results are not all removals; successful application can also retain access.</p><pre>$apply</pre></details><h4>Recorded decision actors</h4>$(ConvertTo-AccessReviewStatusUsersHtml $Summary.Actors -ShowDepartment)"
	'<h4>Individual users and decisions</h4>'
	if (-not @($Summary.Items | Where-Object { $_ }).Count) {
		'<p class="muted">No individual decision records returned.</p>'
		return
	}
	'<p class="muted">Recorded decision-makers are actual actors, not a complete assigned-reviewer roster. An unreviewed item has no recorded actor; automatic/default outcomes can also have no actor.</p>'
	if ($Summary.IncludeManagers) { '<p class="notice">Managers below are current directory relationships, not proof of historical or pending stage assignments. Consult the stage reviewer rule and recorded decision-maker separately.</p>' }
	'<div class="table-wrap"><table class="decision-items"><thead><tr><th>User being reviewed</th><th>Decision</th><th>Recorded decision-maker</th>'
	if ($Summary.IncludeManagers) { '<th>Current line manager (directory snapshot)</th>' }
	'<th>Reviewed (UTC)</th><th>Application result</th></tr></thead><tbody>'
	foreach ($item in $Summary.Items) {
		"<tr><td>$(if ($item.Subject) { ConvertTo-AccessReviewStatusUsersHtml @($item.Subject) } else { 'Subject identity unavailable in the review record.' })</td><td>$(ConvertTo-AccessReviewHtmlText $item.Decision)</td><td>$(if ($item.ReviewedBy) { ConvertTo-AccessReviewStatusUsersHtml @($item.ReviewedBy) -ShowDepartment } else { 'No recorded decision-maker' })</td>"
		if ($Summary.IncludeManagers) {
			"<td>$(if ($item.CurrentManager.User) { ConvertTo-AccessReviewStatusUsersHtml @($item.CurrentManager.User) -ShowDepartment })<small>$(ConvertTo-AccessReviewHtmlText $item.CurrentManager.Note)</small>$(if ($item.CurrentManager.Error) { '<span class="warning">' + (ConvertTo-AccessReviewHtmlText $item.CurrentManager.Error) + '</span>' })$(if (-not $item.CurrentManager) { 'Unavailable / not a user subject' })</td>"
		}
		"<td>$(ConvertTo-AccessReviewHtmlText $item.ReviewedDateTime)</td><td>$(ConvertTo-AccessReviewHtmlText $item.ApplyResult)</td></tr>"
	}
	'</tbody></table></div>'
}

function ConvertTo-AccessReviewStatusChartHtml {
	param([AllowEmptyCollection()][object[]]$Definitions)
	$heading = '<section id="status-dashboard" class="panel"><h2>Reviews by current status</h2><p class="muted">All matched review definitions returned in this snapshot, counted once each. Historical instances and stages are not counted. Inventory filters do not change this chart.</p>'
	if (-not $Definitions.Count) { return "$heading<p>No review definitions available for the status chart.</p></section>" }
	$colors = @{
		Initializing = '#a66340'; NotStarted = '#7b7f88'; Starting = '#9c6aae'
		InProgress = '#b51f3b'; Completing = '#d66c42'; Completed = '#237868'
		AutoReviewing = '#42799a'; AutoReviewed = '#5c628f'; Applying = '#a98118'; Applied = '#386641'
	}
	$groups = @($Definitions | Group-Object {
		if ([string]::IsNullOrWhiteSpace($_.Status)) { 'Unavailable' } else { [string]$_.Status }
	} | Sort-Object Name)
	$stops = [System.Collections.Generic.List[string]]::new()
	$legend = [System.Text.StringBuilder]::new()
	$labels = [System.Collections.Generic.List[string]]::new()
	$position = 0
	$culture = [System.Globalization.CultureInfo]::InvariantCulture
	foreach ($group in $groups) {
		$start = (100.0 * $position / $Definitions.Count).ToString('0.######', $culture)
		$position += $group.Count
		$end = (100.0 * $position / $Definitions.Count).ToString('0.######', $culture)
		$percent = (100.0 * $group.Count / $Definitions.Count).ToString('0.#', $culture)
		$color = if ($colors.ContainsKey($group.Name)) { $colors[$group.Name] } else { '#737373' }
		$stops.Add("$color $start% $end%")
		$labels.Add("$($group.Name): $($group.Count) ($percent%)")
		$name = ConvertTo-AccessReviewHtmlText $group.Name
		$null = $legend.AppendLine("<li data-review-status=`"$name`" data-count=`"$($group.Count)`"><span class=`"status-swatch`" style=`"background:$color`" aria-hidden=`"true`"></span><span class=`"status-label`">$name</span><strong>$($group.Count)</strong><span>$percent%</span></li>")
	}
	$label = ConvertTo-AccessReviewHtmlText ("Reviews by current status. " + ($labels -join '; '))
	"$heading<div class=`"status-dashboard-layout`"><div class=`"status-pie`" role=`"img`" aria-label=`"$label`" style=`"background:conic-gradient($($stops -join ', '))`"></div><ul class=`"status-legend`" aria-label=`"Review status counts and percentages`">$legend</ul></div></section>"
}

function New-AccessReviewStatusHtml {
	param([Parameter(Mandatory)][object]$Report)
	$body = [System.Text.StringBuilder]::new()
	$instances = @($Report.Definitions | ForEach-Object { $_.Instances })
	$known = @($instances | Where-Object { $null -ne $_.Summary })
	$total = [long](($known | ForEach-Object { $_.Summary.Total } | Measure-Object -Sum).Sum)
	$recorded = [long](($known | ForEach-Object { $_.Summary.Recorded } | Measure-Object -Sum).Sum)
	$notReviewed = [long](($known | ForEach-Object { $_.Summary.NotReviewed } | Measure-Object -Sum).Sum)
	$dontKnow = [long](($known | ForEach-Object { $_.Summary.DontKnow } | Measure-Object -Sum).Sum)
	$other = [long](($known | ForEach-Object { $_.Summary.Other } | Measure-Object -Sum).Sum)
	$null = $body.AppendLine('<main><header><div><p class="eyebrow">Identity governance / read-only snapshot</p><h1>Access review status</h1></div><div class="theme-control"><button id="theme" type="button" role="switch" aria-checked="false" aria-label="Dark mode">Dark mode</button><small id="theme-notice" role="status"></small></div></header>')
	$null = $body.AppendLine("<p>Tenant: $(ConvertTo-AccessReviewHtmlText $Report.TenantId)<br>Collected: $(ConvertTo-AccessReviewHtmlText $Report.StartedAt) to $(ConvertTo-AccessReviewHtmlText $Report.FinishedAt)<br>Literal prefixes: $(ConvertTo-AccessReviewHtmlText ($Report.Prefixes -join ', '))</p>")
	$null = $body.AppendLine('<p class="notice">All historical instances returned by Graph. This is a read-only, non-transactional snapshot, not proof of certification or removal. Decision totals below count instance items only, not repeated stage tasks or unique people. Recorded outcomes include Approve, Deny and Don&#39;t know, and may include automatic/default decisions. Names, sign-in addresses, departments and IDs are private.</p>')
	if (-not $Report.Completed -or $Report.Issues.Count) {
		$null = $body.AppendLine("<p class=`"notice warning`" role=`"alert`"><strong>Partial data / collection issues.</strong> $($Report.Issues.Count) issue(s). Unavailable data is not treated as zero. See collection issues below.</p>")
	}
	$null = $body.AppendLine('<section class="metrics" aria-label="Status summary">')
	foreach ($metric in @(
		@('Definitions', $Report.Definitions.Count), @('Historical instances', $instances.Count),
		@('Available instance decision items', $total), @('Recorded outcomes', $recorded),
		@('Not reviewed', $notReviewed), @("Don't know", $dontKnow),
		@('Other / unknown outcomes', $other), @('Collection issues', $Report.Issues.Count)
	)) {
		$null = $body.AppendLine("<div class=`"metric`"><strong>$($metric[1])</strong><span>$(ConvertTo-AccessReviewHtmlText $metric[0])</span></div>")
	}
	$null = $body.AppendLine("</section><p class=`"muted`">Decision-count coverage: $($known.Count) of $($instances.Count) instances. Future/current configured stages and runtime stages are shown separately. Current configuration and directory owners are not historical assignment snapshots.</p>")
	$null = $body.AppendLine((ConvertTo-AccessReviewStatusChartHtml -Definitions @($Report.Definitions)))
	$null = $body.AppendLine('<section class="panel"><h2>Summary by prefix</h2><div class="table-wrap"><table><thead><tr><th>Prefix</th><th>Definitions</th><th>Instances</th><th>Known / all instances</th><th>Items</th><th>Recorded outcomes</th><th>Not reviewed</th></tr></thead><tbody>')
	foreach ($group in @($Report.Definitions | Group-Object Prefix | Sort-Object Name)) {
		$groupInstances = @($group.Group | ForEach-Object { $_.Instances })
		$groupKnown = @($groupInstances | Where-Object { $null -ne $_.Summary })
		$sum = [long](($groupKnown | ForEach-Object { $_.Summary.Total } | Measure-Object -Sum).Sum)
		$done = [long](($groupKnown | ForEach-Object { $_.Summary.Recorded } | Measure-Object -Sum).Sum)
		$pending = [long](($groupKnown | ForEach-Object { $_.Summary.NotReviewed } | Measure-Object -Sum).Sum)
		$null = $body.AppendLine("<tr><td>$(ConvertTo-AccessReviewHtmlText $group.Name)</td><td>$($group.Count)</td><td>$($groupInstances.Count)</td><td>$($groupKnown.Count) / $($groupInstances.Count)</td><td>$sum</td><td>$done</td><td>$pending</td></tr>")
	}
	$null = $body.AppendLine('</tbody></table></div></section><section id="inventory" class="panel"><h2>Review and instance inventory</h2><div class="filters"><label>Search reviews, owners, reviewers or departments<input id="search" type="search"></label><label>Prefix<select id="prefix"><option value="">All prefixes</option>')
	foreach ($prefix in @($Report.Definitions.Prefix | Sort-Object -Unique)) {
		$text = ConvertTo-AccessReviewHtmlText $prefix
		$null = $body.AppendLine("<option value=`"$text`">$text</option>")
	}
	$null = $body.AppendLine('</select></label><label>Instance status<select id="instance-status"><option value="">All instance statuses</option>')
	foreach ($status in @(@($instances.Status) + @('No instances / unavailable') | Where-Object { $_ } | Sort-Object -Unique)) {
		$text = ConvertTo-AccessReviewHtmlText $status
		$null = $body.AppendLine("<option value=`"$text`">$text</option>")
	}
	$null = $body.AppendLine('</select></label><label>Definition status<select id="definition-status"><option value="">All definition statuses</option>')
	foreach ($status in @($Report.Definitions.Status | Where-Object { $_ } | Sort-Object -Unique)) {
		$text = ConvertTo-AccessReviewHtmlText $status
		$null = $body.AppendLine("<option value=`"$text`">$text</option>")
	}
	$null = $body.AppendLine('</select></label></div><p id="inventory-count" role="status"></p><div class="table-wrap"><table><thead><tr><th>Review / prefix</th><th>Definition status</th><th>Instance status</th><th>Instance dates (UTC)</th><th>Recorded / items</th><th>Runtime stages</th></tr></thead><tbody id="status-rows">')
	$details = [System.Text.StringBuilder]::new()
	$detailPages = [System.Collections.Generic.List[string]]::new()
	$index = 0
	foreach ($definition in $Report.Definitions) {
		$entries = @($definition.Instances)
		if (-not $entries.Count) { $entries = @($null) }
		foreach ($instance in $entries) {
			$status = if ($instance) { $instance.Status } else { 'No instances / unavailable' }
			$search = ConvertTo-AccessReviewHtmlText (@{ Review = $definition.DisplayName; Configuration = $definition.Configuration; Owners = $definition.CurrentResources; Instance = $instance } | ConvertTo-Json -Depth 15 -Compress)
			$progress = if ($null -ne $instance.Summary) { "$($instance.Summary.Recorded) / $($instance.Summary.Total)" } else { 'Unavailable / not materialized' }
			$stageCount = if ($instance.StageError) { 'Unavailable' } elseif ($instance) { "$($instance.Stages.Count)" } else { 'Not materialized' }
			$null = $body.AppendLine("<tr data-prefix=`"$(ConvertTo-AccessReviewHtmlText $definition.Prefix)`" data-instance=`"$(ConvertTo-AccessReviewHtmlText $status)`" data-definition=`"$(ConvertTo-AccessReviewHtmlText $definition.Status)`" data-search=`"$search`"><td><button class=`"link view`" data-target=`"status-$index`">$(ConvertTo-AccessReviewHtmlText $definition.DisplayName)</button><small>$(ConvertTo-AccessReviewHtmlText $definition.Prefix)</small></td><td>$(ConvertTo-AccessReviewHtmlText $definition.Status)</td><td><span class=`"badge`">$(ConvertTo-AccessReviewHtmlText $status)</span></td><td>$(ConvertTo-AccessReviewHtmlText $instance.Start)<br>$(ConvertTo-AccessReviewHtmlText $instance.End)</td><td>$progress</td><td>$stageCount</td></tr>")
			$null = $details.AppendLine("<article id=`"status-$index`" class=`"review-detail`" hidden><h3 tabindex=`"-1`">$(ConvertTo-AccessReviewHtmlText $definition.DisplayName)</h3><p>Definition: <code>$(ConvertTo-AccessReviewHtmlText $definition.Id)</code> | $(ConvertTo-AccessReviewHtmlText $definition.Status)<br>Instance: <code>$(ConvertTo-AccessReviewHtmlText $instance.Id)</code> | $(ConvertTo-AccessReviewHtmlText $status)</p><p>$(ConvertTo-AccessReviewHtmlText $definition.Description)</p>")
			if ($definition.Error) { $null = $details.AppendLine("<p class=`"notice warning`">Instance listing unavailable: $(ConvertTo-AccessReviewHtmlText $definition.Error)</p>") }
			$null = $details.AppendLine('<details open><summary>All configured stages (current definition, not historical evidence)</summary><p class="notice">These stages include future stages not returned by the runtime API. Settings may have changed since older instances ran. Named users may be captured owners or fallbacks; their original policy label cannot be inferred from user IDs alone.</p>')
			foreach ($setting in $definition.Configuration) {
				$null = $details.AppendLine("<h4>Configured stage $(ConvertTo-AccessReviewHtmlText $setting.Id)</h4><p>Duration: $(ConvertTo-AccessReviewHtmlText $setting.DurationInDays) days. Depends on: $(ConvertTo-AccessReviewHtmlText ($setting.DependsOn -join ', '))</p>$(ConvertTo-AccessReviewStatusScopesHtml $setting.Reviewers)<h4>Configured fallbacks</h4>$(ConvertTo-AccessReviewStatusScopesHtml $setting.Fallbacks)")
			}
			$null = $details.AppendLine("<details><summary>Current definition resource scope</summary><pre>$(ConvertTo-AccessReviewHtmlText ($definition.Scope | ConvertTo-Json -Depth 10))</pre></details><h4>Current definition resource owners (directory snapshot)</h4>$(ConvertTo-AccessReviewStatusOwnersHtml $definition.CurrentResources)</details>")
			if ($instance) {
				$null = $details.AppendLine("<details><summary>Recorded instance resource scope</summary><pre>$(ConvertTo-AccessReviewHtmlText ($instance.Scope | ConvertTo-Json -Depth 10))</pre></details>")
				$null = $details.AppendLine("<h4>Instance-level outcomes</h4>$(ConvertTo-AccessReviewStatusSummaryHtml $instance.Summary)")
				foreach ($field in @('DecisionError', 'ContactError', 'StageError')) {
					if ($instance.$field) { $null = $details.AppendLine("<p class=`"notice warning`">$(ConvertTo-AccessReviewHtmlText $field): $(ConvertTo-AccessReviewHtmlText $instance.$field)</p>") }
				}
				$null = $details.AppendLine("<details open><summary>Recorded instance reviewers (instance-wide, not mapped to stages)</summary><p>Graph's contactedReviewers collection includes reviewers irrespective of notification delivery. This is not a per-stage roster.</p>$(ConvertTo-AccessReviewStatusUsersHtml $instance.ContactedReviewers -ShowDepartment)</details><h4>Instance reviewer scopes</h4>$(ConvertTo-AccessReviewStatusScopesHtml $instance.Reviewers)<h4>Instance fallbacks</h4>$(ConvertTo-AccessReviewStatusScopesHtml $instance.Fallbacks)<details><summary>Current owners of resources recorded for this instance</summary>$(ConvertTo-AccessReviewStatusOwnersHtml $instance.Resources)</details>")
				$null = $details.AppendLine('<h4>Runtime stages</h4><p class="muted">Sorted by recorded start time and ID; runtime GUIDs are not assumed to map to current configured stage IDs.</p>')
				if (-not $instance.Stages.Count -and -not $instance.StageError) { $null = $details.AppendLine('<p>No runtime stages returned. Instance-level decisions remain separate; no stage count or future stage status is invented.</p>') }
				foreach ($stage in $instance.Stages) {
					$null = $details.AppendLine("<details open class=`"runtime-stage`"><summary>Runtime stage <code>$(ConvertTo-AccessReviewHtmlText $stage.Id)</code> - $(ConvertTo-AccessReviewHtmlText $stage.Status)</summary><p>Start: $(ConvertTo-AccessReviewHtmlText $stage.Start)<br>End: $(ConvertTo-AccessReviewHtmlText $stage.End)</p><h4>Stage reviewer scopes</h4>$(ConvertTo-AccessReviewStatusScopesHtml $stage.Reviewers)<h4>Stage fallbacks</h4>$(ConvertTo-AccessReviewStatusScopesHtml $stage.Fallbacks)$(ConvertTo-AccessReviewStatusSummaryHtml $stage.Summary)")
					if ($stage.Error) { $null = $details.AppendLine("<p class=`"notice warning`">$(ConvertTo-AccessReviewHtmlText $stage.Error)</p>") }
					$null = $details.AppendLine('</details>')
				}
			}
			$null = $details.AppendLine('</article>')
			$detailPages.Add($details.ToString())
			$null = $details.Clear()
			$index++
		}
	}
	$null = $body.AppendLine('</tbody></table></div><div id="pagination"><button id="previous" type="button">Previous</button><span id="page"></span><button id="next" type="button">Next</button></div><noscript>JavaScript enables filtering, paging, theme switching and detail selection. Inventory rows remain visible without JavaScript.</noscript></section><section class="panel" id="review-details"><h2>Stages, reviewers and owners</h2><p id="selection-hint">Choose a review instance above. Search also matches reviewer and owner names, sign-in addresses, departments and IDs.</p>')
	$null = $body.AppendLine('<div id="selected-status-detail"></div>')
	$null = $body.AppendLine('</section><section class="panel"><h2>Collection issues</h2><p>Permission errors, deleted users and unavailable APIs remain visible. Missing data is never proof that no reviewers or owners exist.</p><ul>')
	foreach ($issue in $Report.Issues) {
		$null = $body.AppendLine("<li><strong>$(ConvertTo-AccessReviewHtmlText $issue.Source)</strong>: $(ConvertTo-AccessReviewHtmlText $issue.Message)<small>Definition: $(ConvertTo-AccessReviewHtmlText $issue.DefinitionId) | Instance: $(ConvertTo-AccessReviewHtmlText $issue.InstanceId)</small></li>")
	}
	$null = $body.AppendLine('</ul></section><footer>Read-only report. No reviews, decisions, notifications, owners or memberships were changed. Directory profiles and owners are current lookups, not guaranteed historical identities or assignments.</footer></main>')
	$template = @'
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'">
<title>Access review status</title><style><!--STYLE--></style></head><body><!--REPORT--><script type="application/json" id="status-details"><!--DETAILS--></script><script>
(() => {
/*THEME*/
const rows = [...document.querySelectorAll('#status-rows tr')];
const search = document.getElementById('search');
const prefix = document.getElementById('prefix');
const instanceStatus = document.getElementById('instance-status');
const definitionStatus = document.getElementById('definition-status');
const detailPages = JSON.parse(document.getElementById('status-details').textContent);
const detailHost = document.getElementById('selected-status-detail');
let filtered = rows, current = 0;
function render() {
  const pages = Math.max(1, Math.ceil(filtered.length / 50));
  current = Math.max(0, Math.min(current, pages - 1));
  const visible = new Set(filtered.slice(current * 50, current * 50 + 50));
  for (const row of rows) row.hidden = !visible.has(row);
  document.getElementById('inventory-count').textContent = `${filtered.length} of ${rows.length} inventory entries (one per instance, or a definition without instances).`;
  document.getElementById('page').textContent = `Page ${current + 1} of ${pages}`;
  document.getElementById('previous').disabled = current === 0;
  document.getElementById('next').disabled = current + 1 === pages;
}
function filter() {
  const term = search.value.toLocaleLowerCase();
  filtered = rows.filter(row => (!prefix.value || row.dataset.prefix === prefix.value)
    && (!instanceStatus.value || row.dataset.instance === instanceStatus.value)
    && (!definitionStatus.value || row.dataset.definition === definitionStatus.value)
    && row.dataset.search.toLocaleLowerCase().includes(term));
  current = 0;
  detailHost.replaceChildren();
  document.getElementById('selection-hint').hidden = false;
  render();
}
search.addEventListener('input', filter);
for (const select of [prefix, instanceStatus, definitionStatus]) select.addEventListener('change', filter);
document.getElementById('previous').addEventListener('click', () => { current--; render(); });
document.getElementById('next').addEventListener('click', () => { current++; render(); });
document.addEventListener('click', event => {
  const button = event.target.closest('button.view');
  if (!button) return;
  detailHost.innerHTML = detailPages[Number(button.dataset.target.slice('status-'.length))];
  document.getElementById('selection-hint').hidden = true;
  const detail = document.getElementById(button.dataset.target);
  detail.hidden = false;
  detail.querySelector('h3').focus();
  detail.scrollIntoView({block:'start'});
});
render();
})();
</script></body></html>
'@
	$detailJson = ConvertTo-Json -InputObject @($detailPages.ToArray()) -Compress
	$detailJson = $detailJson.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
	$template.Replace('<!--STYLE-->', (Get-AccessReviewReportStyle)).Replace('/*THEME*/', (Get-AccessReviewReportThemeScript)).Replace('<!--DETAILS-->', $detailJson).Replace('<!--REPORT-->', $body.ToString())
}
