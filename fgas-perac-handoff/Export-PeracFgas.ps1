<#
Excel COM helper for the South Korea PERAC F-gas multiplier fix (see HANDOFF.md).
Run from anywhere; paths resolve relative to this script (repo root = parent folder).

  -Mode Check  : read-only. Exports the five tabs from the UNEDITED workbook with Excel's own
                 CSV writer and compares them with the CSVs in the repo (line endings ignored).
                 Nothing is saved or overwritten.
  -Mode Apply  : sets E37/J37:AI37 and E42/J42:AI42 to the GHG Inventory pegs, runs a full
                 recalculation, exports the five tabs to a temp folder, checks the invariants,
                 and only if everything passes saves the workbook and writes the five CSVs
                 (LF line endings) into InputData/indst/PERAC/. Run only after approval.

Excel must be installed. Close the workbook in Excel first.
#>
param([Parameter(Mandatory)][ValidateSet('Check','Apply','Export')][string]$Mode)
$ErrorActionPreference = 'Stop'
$repo  = Split-Path -Parent $PSScriptRoot
$dir   = Join-Path $repo 'InputData\indst\PERAC'
$book  = Join-Path $dir  'Proc Emis Reductions and Costs.xlsx'
$names = 'fgassubstitution','fgasdestruction','fgasrecovery','chem-inspctmaintretrofit','alum-inspctmaintretrofit'
# Stage exports in TEMP with a unique file prefix. Note: SaveAs must get a plain [string] path;
# passing Join-Path output straight to the COM call makes Excel hang.
$tmp   = $env:TEMP
$tag   = 'peracfgas_' + [guid]::NewGuid().ToString('N').Substring(0,8) + '_'

if (Get-ChildItem $dir -Filter '~$*' -Force -ErrorAction SilentlyContinue) {
    throw "Lock file found in $dir - someone has a PERAC workbook open. Stop and reconcile first."
}

function Read-Bytes($p) { [IO.File]::ReadAllBytes($p) }
function To-LF([byte[]]$b) {                      # drop every CR that precedes an LF
    $o = New-Object 'System.Collections.Generic.List[byte]' $b.Length
    for ($i=0; $i -lt $b.Length; $i++) { if ($b[$i] -eq 13 -and $i+1 -lt $b.Length -and $b[$i+1] -eq 10) { continue }; $o.Add($b[$i]) }
    ,$o.ToArray()
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($book, 0, ($Mode -ne 'Apply'))   # UpdateLinks=0, ReadOnly for Check
    if ($Mode -eq 'Apply' -and $wb.ReadOnly) { throw 'Workbook opened read-only (open elsewhere). Stop and reconcile.' }

    if ($Mode -eq 'Apply') {
        $m = $wb.Worksheets.Item('Multipliers and Adjustments')
        foreach ($p in @(@{row=37;inv='$F$26'}, @{row=42;inv='$F$31'})) {
            $r = $p.row
            $m.Range("E$r").Formula = "='GHG Inventory'!$($p.inv)"
            $m.Range("J${r}:AI$r").Formula = "=`$E$r"
        }
        $xl.CalculateFull()
        $e37 = $m.Range('E37').Value2; $e42 = $m.Range('E42').Value2
        if ([math]::Abs($e37-0.59125) -gt 1e-5 -or [math]::Abs($e42-0.69486) -gt 1e-5) { throw "Unexpected pegs E37=$e37 E42=$e42" }
        foreach ($c in 'J37','AI37','J42','AI42') { if ([math]::Abs($m.Range($c).Value2 - $(if ($c -like '*37') {$e37} else {$e42})) -gt 1e-12) { throw "$c does not equal its peg" } }
    }

    foreach ($n in $names) {                       # export each tab as values via Excel's CSV writer
        $ws = $wb.Worksheets.Item("PERAC-$n")
        $ws.Copy()
        $nb = $xl.ActiveWorkbook
        $u = $nb.Worksheets.Item(1).UsedRange; $u.Value2 = $u.Value2
        $nb.SaveAs([string](Join-Path $tmp "${tag}PERAC-$n.csv"), [int]6)   # xlCSV
        $nb.Close($false)
    }

    $ok = $true
    foreach ($n in $names) {
        $new = To-LF (Read-Bytes (Join-Path $tmp "${tag}PERAC-$n.csv"))
        $old = To-LF (Read-Bytes (Join-Path $dir  "PERAC-$n.csv"))
        if ($Mode -eq 'Check') {
            $same = ($new.Length -eq $old.Length) -and -not (Compare-Object $new $old -SyncWindow 0 | Select-Object -First 1)
            '{0,-28} {1}' -f $n, $(if ($same) {'IDENTICAL to current CSV'} else {'DIFFERS from current CSV'})
            if (-not $same) { $ok = $false }
        } else {
            $lines = [Text.Encoding]::ASCII.GetString($new).TrimEnd("`n").Split("`n")
            $oldCols = ([Text.Encoding]::ASCII.GetString($old).Split("`n")[0]).Split(',').Count
            $bad = @()
            if ($lines.Count -ne 74) { $bad += "rows=$($lines.Count)" }
            if ($lines[0].Split(',').Count -ne $oldCols) { $bad += 'column count changed' }
            foreach ($l in $lines[2..($lines.Count-1)]) { $f = $l.Split(','); if ([double]$f[2] -ne 0) { $bad += '2020 nonzero'; break } }
            foreach ($l in $lines[2..($lines.Count-1)]) { if (($l.Split(',')[2..32] | ForEach-Object {[double]$_} | Where-Object {$_ -lt 0}).Count) { $bad += 'negative value'; break } }
            if ($bad) { $ok = $false; "$n FAILED: $($bad -join '; ')" } else { "$n ok" }
            [IO.File]::WriteAllBytes((Join-Path $tmp "${tag}LF-$n.csv"), $new)
        }
    }

    if ($Mode -eq 'Apply') {
        if (-not $ok) { throw 'Invariant failure - workbook NOT saved, CSVs NOT written.' }
        $wb.Save()
        foreach ($n in $names) { Copy-Item (Join-Path $tmp "${tag}LF-$n.csv") (Join-Path $dir "PERAC-$n.csv") -Force }
        'Workbook saved and five CSVs written (LF). Next: run the gates in HANDOFF.md.'
    } elseif ($Mode -eq 'Export') {
        if (-not $ok) { throw 'Invariant failure - CSVs NOT written.' }
        foreach ($n in ($names | Where-Object { $_ -ne 'alum-inspctmaintretrofit' })) { Copy-Item (Join-Path $tmp "${tag}LF-$n.csv") (Join-Path $dir "PERAC-$n.csv") -Force }
        'Four F-gas CSVs written (LF); alum-IMR left untouched. Workbook not saved in this mode.'
    } elseif (-not $ok) { 'Check FAILED: workbook and CSVs are out of sync. Stop and ask.' } else { 'Check passed.' }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item (Join-Path $tmp "$tag*") -Force -ErrorAction SilentlyContinue
}
