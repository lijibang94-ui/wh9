param(
    [Parameter(Mandatory = $true)]
    [string]$IniPath,

    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

if (-not (Test-Path -LiteralPath $IniPath)) {
    throw "INI file not found: $IniPath"
}

$lines = Get-Content -LiteralPath $IniPath -Encoding UTF8
$section = ""
$date = ""
$time = ""
$rows = New-Object System.Collections.Generic.List[object]

foreach ($line in $lines) {
    $trim = $line.Trim()
    if ($trim -match '^\[(.+)\]$') {
        $section = $Matches[1]
        continue
    }

    if ($trim -eq "" -or $trim.StartsWith(";")) {
        continue
    }

    $pair = $trim -split '=', 2
    if ($pair.Count -ne 2) {
        continue
    }

    $key = $pair[0]
    $value = $pair[1]

    if ($section -eq "Meta") {
        if ($key -eq "Date") { $date = $value }
        if ($key -eq "Time") { $time = $value }
        continue
    }

    if ($section -eq "RiseRank" -or $section -eq "OpiRank") {
        $fields = @{
            RankType = $section
            RankNo = $key
            Date = $date
            Time = $time
            Contract = ""
            RisePct = ""
            DayDeltaOpi = ""
            DayDeltaOpiPct = ""
            Capital = ""
        }

        $parts = $value -split ','
        if ($parts.Count -gt 0) {
            $fields.Contract = $parts[0]
        }

        for ($i = 1; $i -lt $parts.Count; $i++) {
            $kv = $parts[$i] -split '=', 2
            if ($kv.Count -eq 2 -and $fields.ContainsKey($kv[0])) {
                $fields[$kv[0]] = $kv[1]
            }
        }

        $rows.Add([pscustomobject]$fields)
    }
}

$rows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Output "written: $CsvPath"
