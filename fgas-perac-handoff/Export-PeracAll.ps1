<#
Export every PERAC-* tab (except PERAC-MCD) from the PERAC workbook as values to LF CSVs in a staging folder.
Read-only on the workbook. Output folder: $env:TEMP\peracall_stage (printed). Then run
  python fgas-perac-handoff\check_korea_dup_gates.py install   # copies only tabs whose values changed (>1e-5 relative)
  python fgas-perac-handoff\check_korea_dup_gates.py check
SaveAs gets a plain [string] path in TEMP root (a Join-Path result or a new subfolder makes Excel hang).
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$dir  = Join-Path $repo 'InputData\indst\PERAC'
$book = Join-Path $dir 'Proc Emis Reductions and Costs.xlsx'
$stage = Join-Path $env:TEMP 'peracall_stage'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage | Out-Null
if (Get-ChildItem $dir -Filter '~$*' -Force -ErrorAction SilentlyContinue) { throw "Lock file in $dir - a workbook is open. Stop and reconcile." }
$tag = 'peracall_' + [guid]::NewGuid().ToString('N').Substring(0,8) + '_'

function To-LF([byte[]]$b) {
    $o = New-Object 'System.Collections.Generic.List[byte]' $b.Length
    for ($i=0; $i -lt $b.Length; $i++) { if ($b[$i] -eq 13 -and $i+1 -lt $b.Length -and $b[$i+1] -eq 10) { continue }; $o.Add($b[$i]) }
    ,$o.ToArray()
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open([string]$book, 0, $true)
    $names = @(); foreach ($s in $wb.Worksheets) { if ($s.Name -like 'PERAC-*' -and $s.Name -ne 'PERAC-MCD') { $names += $s.Name.Substring(6) } }
    "tabs: $($names.Count)"
    foreach ($n in $names) {
        $ws = $wb.Worksheets.Item("PERAC-$n"); $ws.Copy()
        $nb = $xl.ActiveWorkbook
        $u = $nb.Worksheets.Item(1).UsedRange; $u.Value2 = $u.Value2
        $nb.SaveAs([string](Join-Path $env:TEMP "${tag}PERAC-$n.csv"), [int]6)
        $nb.Close($false)
        $new = To-LF ([IO.File]::ReadAllBytes((Join-Path $env:TEMP "${tag}PERAC-$n.csv")))
        $lines = [Text.Encoding]::ASCII.GetString($new).TrimEnd("`n").Split("`n")
        $oldCols = ([Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes((Join-Path $dir "PERAC-$n.csv"))).Split("`n")[0]).Split(',').Count
        $bad = @()
        if ($lines.Count -ne 74) { $bad += "rows=$($lines.Count)" }
        if ($lines[0].Split(',').Count -ne $oldCols) { $bad += 'column count changed' }
        foreach ($l in $lines[2..($lines.Count-1)]) { $f = $l.Split(','); if ([double]$f[2] -ne 0) { $bad += '2020 nonzero'; break } }
        foreach ($l in $lines[2..($lines.Count-1)]) { if (($l.Split(',')[2..32] | ForEach-Object {[double]$_} | Where-Object {$_ -lt 0}).Count) { $bad += 'negative value'; break } }
        if ($bad) { throw "$n FAILED: $($bad -join '; ')" }
        [IO.File]::WriteAllBytes((Join-Path $stage "PERAC-$n.csv"), $new); "$n ok"
    }
    "staged in $stage"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item (Join-Path $env:TEMP "$tag*") -Force -ErrorAction SilentlyContinue
}
