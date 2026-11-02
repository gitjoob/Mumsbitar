function Write-AccessReviewLog {
	param(
		[Parameter(Mandatory)][hashtable]$Log,
		[Parameter(Mandatory)][hashtable]$Entry
	)

	$record = [ordered]@{
		Timestamp = [datetime]::UtcNow.ToString('o')
		RunId = $Log.RunId
	}
	foreach ($key in $Entry.Keys) { $record[$key] = $Entry[$key] }
	try {
		$null = [System.IO.Directory]::CreateDirectory((Split-Path -Parent $Log.Path))
		[System.IO.File]::AppendAllText(
			$Log.Path,
			(($record | ConvertTo-Json -Depth 10 -Compress) + [Environment]::NewLine),
			[System.Text.UTF8Encoding]::new($false)
		)
	}
	catch {
		throw "Unable to write creation log '$($Log.Path)': $($_.Exception.Message). Check any review IDs already returned before rerunning."
	}
}

function Start-AccessReviewLog {
	param(
		[Parameter(Mandatory)][string]$ScriptPath,
		[Parameter(Mandatory)][hashtable]$Parameters,
		[Parameter(Mandatory)][hashtable]$Settings
	)

	$log = @{
		RunId = [guid]::NewGuid().ToString()
		Path = Join-Path (Split-Path -Parent $ScriptPath) "logs/$([System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)).jsonl"
	}
	$syntaxParts = @("& '$($ScriptPath.Replace("'", "''"))'")
	foreach ($key in @($Parameters.Keys | Sort-Object)) {
		$value = $Parameters[$key]
		if ($value -is [System.Management.Automation.SwitchParameter]) {
			$syntaxParts += "-${key}:`$$($value.IsPresent.ToString().ToLowerInvariant())"
		}
		elseif ($value -is [array]) {
			$items = @($value | ForEach-Object { "'$(([string]$_).Replace("'", "''"))'" })
			$syntaxParts += "-$key @($($items -join ', '))"
		}
		else {
			$text = if ($value -is [datetime]) { $value.ToString('o') } else { [string]$value }
			$syntaxParts += "-$key '$($text.Replace("'", "''"))'"
		}
	}
	$entry = @{
		Event = 'RunStarted'
		Syntax = $syntaxParts -join ' '
		WorkingDirectory = (Get-Location).Path
	}
	foreach ($key in $Settings.Keys) { $entry[$key] = $Settings[$key] }
	Write-AccessReviewLog -Log $log -Entry $entry
	$log
}
