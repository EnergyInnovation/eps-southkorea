<#
Korea EPA duplicate-copy correction (all EPA-derived PERAC blocks).

Why: EPA's combined MACC file lists Korea twice ('South Korea' and 'Korea, South', identical rows). macc_cleaner.py
filters on country_code only and bucketer.py sums quantities, so every EPA_MACCs quantity for KOR is exactly 2x the
EPA annex. This script does NOT overwrite the raw quantities (column G). It adds a visible per-country correction
table on 'Multipliers and Adjustments' (KOR = 0.5), applies it in EPA_MACCs column I (which already carries the
F-gas source multiplier), and points every remaining MACC Anchors SUMIFS at column I.

Run with both workbooks closed in Excel (only the PERAC workbook is opened):
  powershell -ExecutionPolicy Bypass -File fgas-perac-handoff\Apply-KoreaDupCorrection.ps1
Then run Export-PeracAll.ps1 and check_korea_dup_gates.py check.
The workbook is saved only if every sanity check passes.
#>
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$perac = Join-Path $repo 'InputData\indst\PERAC\Proc Emis Reductions and Costs.xlsx'
if (Get-ChildItem (Split-Path $perac) -Filter '~$*' -Force -ErrorAction SilentlyContinue) { throw 'Lock file in PERAC folder: the workbook is open somewhere. Stop and reconcile.' }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $pw = $xl.Workbooks.Open([string]$perac, 0, $false); if ($pw.ReadOnly) { throw 'PERAC opened read-only.' }
    $origCalc = $xl.Calculation; $xl.Calculation = -4135
    $mm  = $pw.Worksheets.Item('Multipliers and Adjustments')
    $em  = $pw.Worksheets.Item('EPA_MACCs')
    $mac = $pw.Worksheets.Item('MACC Anchors')
    $ab  = $pw.Worksheets.Item('About')

    # ---- preconditions: step-2 state present, correction not yet applied
    if (-not $em.Range('I2').HasFormula) { throw 'EPA_MACCs!I2 has no formula: apply Apply-SourcePeg.ps1 first.' }
    if ([string]$em.Range('I2').Formula -like '*$E$513*') { throw 'Correction already applied.' }
    if ($mm.Range('C512').Value2 -or $mm.Range('C513').Value2) { throw 'Multipliers and Adjustments rows 512-516 are not empty.' }
    if ($ab.Range('A89').Value2) { throw 'About!A89 is not empty.' }
    $last = $em.UsedRange.Rows.Count

    # ---- before totals (KOR, raw) for the sanity check
    $before = $xl.Evaluate('SUMIFS(EPA_MACCs!$G$2:$G$' + $last + ',EPA_MACCs!$D$2:$D$' + $last + ',"KOR")')

    # ---- correction table
    $mm.Range('C512').Value2 = 'EPA duplicate-copy correction by country. EPA combined MACC file lists Korea twice (South Korea and Korea, South; identical rows) and macc_cleaner.py / bucketer.py add both copies, so every KOR quantity in EPA_MACCs column G is 2x the EPA annex. EPA_MACCs column I multiplies column G by the factor below (blank/other countries = 1) and all MACC Anchors blocks read column I. Remove the factor once the builder drops the duplicate rows.'
    $mm.Range('C513').Value2 = 'KOR'; $mm.Range('D513').Value2 = 'quantity correction'; $mm.Range('E513').Value2 = 0.5
    $mm.Range('C514').Value2 = '(spare: iso3 in column C, factor in column E)'

    # ---- column I: add the correction factor to the existing formula
    $ma = "'Multipliers and Adjustments'!"
    $old = [string]$em.Range('I2').Formula
    $fI  = $old + '*IFERROR(INDEX(' + $ma + '$E$513:$E$516,MATCH(D2,' + $ma + '$C$513:$C$516,0)),1)'
    $em.Range('I1').Formula = [string]('quantity (Mt CO2e) used by MACC Anchors = raw quantity x duplicate-copy correction (Multipliers and Adjustments E513) x F-gas source multiplier for the four F-gas policies (rows 507-510)')
    $em.Range([string]('I2:I' + $last)).Formula = [string]$fI

    # ---- every remaining MACC Anchors SUMIFS reads column I
    $nG0 = $xl.Evaluate('SUMPRODUCT(--ISNUMBER(SEARCH("EPA_MACCs!$G:$G",FORMULATEXT(''MACC Anchors''!A1:K1200))))')
    [void]$mac.UsedRange.Replace('EPA_MACCs!$G:$G','EPA_MACCs!$I:$I',2,1,$true)

    $ab.Range('A89').Formula = [string]('2026-11-16 (Korea) - EPA duplicate-copy correction: EPA combined MACC file lists Korea twice (South Korea and Korea, South) and the builder adds both, so all KOR quantities in EPA_MACCs column G are 2x the EPA annex. Added a visible factor (Multipliers and Adjustments E513 = 0.5), applied in EPA_MACCs column I, and every MACC Anchors block now reads column I (the F-gas blocks already did). Raw column G is unchanged. All EPA-derived Korea curves roughly halve; cement and tillage are unaffected.')

    $xl.Calculate()
    $xl.Calculation = $origCalc

    # ---- sanity checks (nothing saved if any fails)
    $errI = $xl.Evaluate('SUMPRODUCT(--ISERROR(EPA_MACCs!I2:I' + $last + '))')
    $nG1 = $xl.Evaluate('SUMPRODUCT(--ISNUMBER(SEARCH("EPA_MACCs!$G:$G",FORMULATEXT(''MACC Anchors''!A1:K1200))))')
    $nI  = $xl.Evaluate('SUMPRODUCT(--ISNUMBER(SEARCH("EPA_MACCs!$I:$I",FORMULATEXT(''MACC Anchors''!A1:K1200))))')
    $afterNonF = $xl.Evaluate('SUMPRODUCT((EPA_MACCs!$D$2:$D$' + $last + '="KOR")*(ISNUMBER(MATCH(EPA_MACCs!$H$2:$H$' + $last + ',{"F-gas destruction","F-gas recovery and recycling","F-gas substitution","chemicals - F-gas inspection maintenance retrofitting"},0))=FALSE)*EPA_MACCs!$I$2:$I$' + $last + ')')
    $rawNonF   = $xl.Evaluate('SUMPRODUCT((EPA_MACCs!$D$2:$D$' + $last + '="KOR")*(ISNUMBER(MATCH(EPA_MACCs!$H$2:$H$' + $last + ',{"F-gas destruction","F-gas recovery and recycling","F-gas substitution","chemicals - F-gas inspection maintenance retrofitting"},0))=FALSE)*EPA_MACCs!$G$2:$G$' + $last + ')')
    "errors in EPA_MACCs!I: $errI; MACC Anchors cells reading column G before/after: $nG0 / $nG1 (expect 0 after); reading column I after: $nI"
    "KOR non-F-gas, I / G = {0:N6} (expect 0.5)" -f ($afterNonF / $rawNonF)
    if ($errI -ne 0) { throw 'Errors in EPA_MACCs!I - nothing saved.' }
    if ($nG1 -ne 0) { throw 'MACC Anchors cells still read column G - nothing saved.' }
    if ([math]::Abs($afterNonF / $rawNonF - 0.5) -gt 1e-9) { throw 'Correction not exactly 0.5 for non-F-gas KOR rows - nothing saved.' }
    $pw.Save()
    'Saved PERAC workbook. Next: Export-PeracAll.ps1, then check_korea_dup_gates.py check.'
}
finally {
    if ($pw) { $pw.Close($false) }
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
