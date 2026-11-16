<#
Step 2 of the consolidated South Korea F-gas handoff, version 2: a TIME-SERIES per-source multiplier that is
computed in BPE and read into the PERAC 'Multipliers and Adjustments' tab.

Run on the step-1 PERAC workbook, with both workbooks closed in Excel:
  powershell -ExecutionPolicy Bypass -File fgas-perac-handoff\Apply-SourcePeg.ps1
Then export the CSVs (Export-PeracFgas.ps1 -Mode Export) and run check_source_peg_gates.py check.

BPE (indst/BPE/BAU Process Emissions.xlsx): new tab 'F-gas Source Multipliers'. For each source group
(Electronics, ODS substitutes, Electrical equipment SF6, HFC-23 by-product), every year 2020-2050:
   multiplier(t) = BAU F-gases of the group(t) / EPA Korea baseline of the group's sources(t)
BAU(t) is read from 'BPE-F-gases_pre' (industries 26, 27, 31T33); the EPA baseline is a live SUMIFS on
'EPA non-CO2 Data'. HFC-23 (HCFC-22 production) = 0 because Korea reports no HCFC-22 production.

PERAC: the four multiplier series are pasted as values into a new block at the bottom of
'Multipliers and Adjustments' (same columns/years as the existing matrix); a small 'F-gas Source Map' tab maps
EPA_MACCs source codes to groups; EPA_MACCs column I multiplies each F-gas policy row by the multiplier for its
group and year; the four F-gas blocks of 'MACC Anchors' (rows 307-608) read column I; row 37 is set to 1
(J37:AI37 = $E37) so it cannot stack. Neither workbook is saved unless every sanity check passes.
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$perac = Join-Path $repo 'InputData\indst\PERAC\Proc Emis Reductions and Costs.xlsx'
$bpe   = Join-Path $repo 'InputData\indst\BPE\BAU Process Emissions.xlsx'
foreach ($f in $perac, $bpe) { if (Get-ChildItem (Split-Path $f) -Filter '~$*' -Force -ErrorAction SilentlyContinue) { throw "Lock file in $(Split-Path $f): a workbook is open somewhere. Stop and reconcile." } }

$groups = @(
  @('Electronics',               'computers and electronics 26'),
  @('ODS substitutes',           'appliances and electrical equipment 27'),
  @('Electrical equipment SF6',  'other manufacturing 31T33'),
  @('HFC-23 by-product',         ''))
$map = @(
  @('SEMI','Semiconductor Manufacturing','Electronics'), @('FPD','Flat Panel Display Manufacturing','Electronics'), @('PV','Photovoltaic Cell Manufacturing','Electronics'),
  @('RAC','Refrigeration and Air Conditioning','ODS substitutes'), @('FOAM','Foams Mfg, Use, and Disposal','ODS substitutes'), @('AERO','Aerosols','ODS substitutes'),
  @('FIRE','Fire Protection','ODS substitutes'), @('SOLV','Solvent Use','ODS substitutes'),
  @('EPS','Electric Power Systems','Electrical equipment SF6'), @('HCFC','HCFC-22 Production','HFC-23 by-product'))

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $bw = $xl.Workbooks.Open([string]$bpe, 0, $false);   if ($bw.ReadOnly) { throw 'BPE opened read-only.' }
    $pw = $xl.Workbooks.Open([string]$perac, 0, $false); if ($pw.ReadOnly) { throw 'PERAC opened read-only.' }
    $origCalc = $xl.Calculation; $xl.Calculation = -4135

    foreach ($s in $bw.Worksheets) { if ($s.Name -eq 'F-gas Source Multipliers') { throw 'BPE tab F-gas Source Multipliers already exists.' } }
    foreach ($s in $pw.Worksheets) { if ($s.Name -eq 'F-gas Source Map') { throw 'PERAC tab F-gas Source Map already exists - step 2 may already be applied.' } }
    $mm = $pw.Worksheets.Item('Multipliers and Adjustments')
    if ($mm.Range('J37').HasFormula -or [math]::Abs($mm.Range('E37').Value2 - 1) -lt 1e-9) { throw 'PERAC row 37 is not in the step-1 state (hard-coded growth values). Start from the step-1 workbook.' }

    # ================= BPE: new tab =================
    $bs = $bw.Worksheets.Add([Type]::Missing, $bw.Worksheets.Item('Multiplier F-gases')); $bs.Name = 'F-gas Source Multipliers'
    $bs.Range('A1').Value2 = 'F-gas source multipliers by year = BAU F-gases of the source group / EPA Korea baseline of the same sources (feeds PERAC Multipliers and Adjustments)'
    $bs.Range('A2').Value2 = "Per year t: multiplier(t) = BAU(t) / EPA(t). BAU(t) is the group's industry row in BPE-F-gases_pre (Mt CO2e = g / 10^12). EPA(t) sums 'EPA non-CO2 Data' (EPA 2025 edition, Industrial Processes, MMT CO2e) over the sources mapped to the group below, for the country on 'Country Selector'. HFC-23 by-product (HCFC-22 production) = 0 because Korea reports no HCFC-22 production. The PERAC workbook pastes rows 16-19 as values into Multipliers and Adjustments; re-paste when this tab changes."
    $bs.Range('A4').Value2 = 'BAU F-gases (Mt CO2e)'; $bs.Range('B4').Value2 = 'Industry row in BPE-F-gases_pre'
    for ($y = 2020; $y -le 2050; $y++) { $bs.Cells.Item(4, 3 + $y - 2020).Value2 = $y }
    for ($i = 0; $i -lt 4; $i++) { $rr = 5 + $i; $bs.Cells.Item($rr,1).Value2 = [string]$groups[$i][0]; if ($groups[$i][1]) { $bs.Cells.Item($rr,2).Value2 = [string]$groups[$i][1] } }
    $bs.Range('B8').Value2 = 'none (Korea reports no HCFC-22 production)'
    $bs.Range('C5:AG7').Formula = '=INDEX(''BPE-F-gases_pre''!$B$3:$AF$45,MATCH($B5,''BPE-F-gases_pre''!$A$3:$A$45,0),MATCH(C$4,''BPE-F-gases_pre''!$B$2:$AF$2,0))/10^12'
    $bs.Range('C8:AG8').Formula = '=0'
    $bs.Range('A9').Value2 = 'EPA Korea baseline (Mt CO2e)'
    for ($i = 0; $i -lt 4; $i++) { $rr = 10 + $i; $bs.Cells.Item($rr,1).Value2 = [string]$groups[$i][0] }
    $bs.Range('C10:AG13').Formula = '=SUMPRODUCT(($C$23:$C$32=$A10)*SUMIFS(''EPA non-CO2 Data''!$G:$G,''EPA non-CO2 Data''!$A:$A,''Country Selector''!$A$2,''EPA non-CO2 Data''!$B:$B,"Industrial Processes",''EPA non-CO2 Data''!$C:$C,$B$23:$B$32,''EPA non-CO2 Data''!$F:$F,C$4))'
    $bs.Range('A15').Value2 = 'Multiplier = BAU / EPA baseline (pasted into PERAC Multipliers and Adjustments)'
    for ($i = 0; $i -lt 4; $i++) { $rr = 16 + $i; $bs.Cells.Item($rr,1).Value2 = [string]$groups[$i][0] }
    $bs.Range('C16:AG18').Formula = '=IF(C10=0,0,C5/C10)'
    $bs.Range('C19:AG19').Formula = '=0'
    $bs.Range('A22').Value2 = 'EPA source code (EPA_MACCs)'; $bs.Range('B22').Value2 = 'EPA baseline source name'; $bs.Range('C22').Value2 = 'Source group'
    for ($i = 0; $i -lt 10; $i++) { for ($j = 0; $j -lt 3; $j++) { $rr = 23 + $i; $cc = 1 + $j; $bs.Cells.Item($rr,$cc).Value2 = [string]$map[$i][$j] } }
    $bs.Columns.Item(1).ColumnWidth = 32; $bs.Columns.Item(2).ColumnWidth = 40
    $bs.Calculate()

    function Count-Errors($ws, $addr) { $vals = $ws.Range($addr).Value2; $n = 0; foreach ($v in $vals) { if ($v -is [int] -and $v -lt -2146826000) { $n++ } }; $n }
    $errB = Count-Errors $bs 'A1:AG32'
    if ($errB -ne 0) { throw "BPE tab has $errB error cells - nothing saved." }
    foreach ($r in 16..18) { foreach ($c in 3..33) { $v = $bs.Cells.Item($r,$c).Value2; if ($v -le 0 -or $v -gt 3) { throw "BPE multiplier out of range at row $r col $c = $v - nothing saved." } } }
    'BPE multipliers (2020 / 2021 / 2025 / 2030 / 2035 / 2040 / 2050):'
    foreach ($r in 16..19) { '  {0,-26} {1}' -f $bs.Cells.Item($r,1).Value2, ((0,1,5,10,15,20,30 | ForEach-Object { '{0:N3}' -f $bs.Cells.Item($r,3+$_).Value2 }) -join '  ') }

    # ================= PERAC: read the series =================
    $lastRow = $mm.Cells.Item($mm.Rows.Count, 3).End(-4162).Row
    $r0 = $lastRow + 2
    $mm.Cells.Item($r0, 3).Value2 = 'F-gas source multipliers = BAU F-gases of the source group(t) / EPA Korea baseline(t). Pasted values from BPE tab F-gas Source Multipliers rows 16-19 (same years as row 22); re-paste when BPE changes. Read by EPA_MACCs column I.'
    for ($i = 0; $i -lt 4; $i++) {
        $r = $r0 + 1 + $i
        $mm.Cells.Item($r,3).Value2 = [string]$groups[$i][0]; $mm.Cells.Item($r,4).Value2 = 'F-gases'
        for ($c = 5; $c -le 35; $c++) {
            $y = [int]$mm.Cells.Item(22,$c).Value2
            if ($y -lt 2020 -or $y -gt 2050) { throw "Unexpected year $y in row 22 col $c" }
            $srow = 16 + $i; $scol = 3 + $y - 2020; $val = [double]$bs.Cells.Item($srow,$scol).Value2
            try { $mm.Cells.Item($r,$c).Formula = $val.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } catch { throw ('paste failed at row ' + $r + ' col ' + $c + ' val ' + $val + ' (' + $val.GetType().Name + ') lastRow ' + $lastRow + ' r0 ' + $r0 + ': ' + $_) }
        }
    }
    $rA = $r0 + 1; $rB = $r0 + 4

    # --- PERAC source map tab ---
    $wm = $pw.Worksheets.Add([Type]::Missing, $pw.Worksheets.Item('Multipliers and Adjustments')); $wm.Name = 'F-gas Source Map'
    $wm.Range('A1').Formula = [string]('EPA_MACCs source code to F-gas source group (groups match Multipliers and Adjustments rows ' + $rA + '-' + $rB + ')')
    $wm.Range('A3').Value2 = 'EPA_MACCs source code'; $wm.Range('B3').Value2 = 'Source group'
    for ($i = 0; $i -lt 10; $i++) { $rr = 4 + $i; $wm.Cells.Item($rr,1).Value2 = [string]$map[$i][0]; $wm.Cells.Item($rr,2).Value2 = [string]$map[$i][2] }
    $wm.Range('A15').Value2 = 'A source code that appears under an F-gas policy but is missing here returns #N/A in EPA_MACCs column I on purpose. MG is handled by row 42 and the other-metals curve.'

    # --- EPA_MACCs helper column I ---
    $em = $pw.Worksheets.Item('EPA_MACCs'); $last = $em.UsedRange.Rows.Count
    if ($em.Range('I1').Value2) { throw 'EPA_MACCs!I1 is not empty.' }
    $em.Range('I1').Formula = [string]('quantity (Mt CO2e) after F-gas source multiplier (year-specific; see Multipliers and Adjustments rows ' + $rA + '-' + $rB + ')')
    $ma = "'Multipliers and Adjustments'!"
    # H holds 91 pre-existing #N/A constants (EUR LIVE rows); ISNA leaves those rows at their raw quantity
    $fI = '=IF(ISNA(H2),G2,IF(OR(H2="F-gas destruction",H2="F-gas recovery and recycling",H2="F-gas substitution",H2="chemicals - F-gas inspection maintenance retrofitting"),G2*INDEX(' + $ma + '$E$' + $rA + ':$AI$' + $rB + ',MATCH(VLOOKUP(B2,''F-gas Source Map''!$A$4:$B$13,2,FALSE),' + $ma + '$C$' + $rA + ':$C$' + $rB + ',0),MATCH(C2,' + $ma + '$E$22:$AI$22,0)),G2))'
    $addr = [string]('I2:I' + $last); $em.Range($addr).Formula = [string]$fI

    # --- MACC Anchors blocks read column I ---
    $mac = $pw.Worksheets.Item('MACC Anchors')
    [void]$mac.Range('A307:K608').Replace('EPA_MACCs!$G:$G','EPA_MACCs!$I:$I',2,1,$true)

    # --- row 37 neutralised ---
    $mm.Range('E37').Value2 = 1
    $mm.Range('J37:AI37').Formula = '=$E37'
    [void]$mm.Range('E37').AddComment([string]('F-gas curves are scaled per source group and year from rows ' + $rA + '-' + $rB + ' of this tab (EPA_MACCs column I); row 37 is kept at 1 so the two never stack. Do not set it back to the GHG Inventory F26 peg (0.591): that would scale the curves twice.'))

    $ab = $pw.Worksheets.Item('About')
    if ($ab.Range('A88').Value2) { throw 'About!A88 is not empty.' }
    $ab.Range('A88').Formula = [string]('2026-11-16 (Korea) - F-gas curves scaled per source group and year: BPE tab F-gas Source Multipliers computes multiplier(t) = BAU F-gases of the group / EPA Korea baseline of its sources; the series is pasted into Multipliers and Adjustments rows ' + $rA + '-' + $rB + '; EPA_MACCs column I applies it to the four F-gas policies (new tab F-gas Source Map maps source codes to groups); the F-gas blocks of MACC Anchors read column I; HFC-23 destruction (HCFC) = 0; row 37 set to 1 (J37:AI37 = $E37) so it cannot stack. Replaces the pasted NF3 growth-index values (1.11 -> 3.05).')

    $xl.Calculate()   # recalculates only dirty cells (our edits and their dependents) in both open workbooks; BPE's other cached values stay untouched
    $xl.Calculation = $origCalc
    $errI = $xl.Evaluate('SUMPRODUCT(--ISERROR(EPA_MACCs!I2:I' + $last + '))')
    $nRep = $xl.Evaluate('SUMPRODUCT(--ISNUMBER(SEARCH("EPA_MACCs!$I:$I",FORMULATEXT(''MACC Anchors''!C309:I380))))')
    "errors in EPA_MACCs!I: $errI; substitution-block cells reading column I: $nRep (expect 504); PERAC block rows ${rA}-${rB}"
    if ($errI -ne 0) { throw 'Errors in EPA_MACCs!I - nothing saved.' }
    if ([math]::Abs($mm.Range('AI37').Value2 - 1) -gt 1e-12) { throw 'Row 37 not neutral - nothing saved.' }
    $errB2 = Count-Errors $bs 'A1:AG32'
    if ($errB2 -ne 0) { throw 'BPE errors after full recalculation - nothing saved.' }
    $pw.Save(); $bw.Save()
    'Saved PERAC and BPE. Next: Export-PeracFgas.ps1 -Mode Export, then check_source_peg_gates.py check.'
}
finally {
    if ($bw) { $bw.Close($false) }
    if ($pw) { $pw.Close($false) }
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
