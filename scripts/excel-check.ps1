# Opens the scenario exports in a private, hidden Excel, recalculates, and checks them against
# the unedited fixture: Excel opens them without repair, no new formula errors, the typed values
# read back, cleared cells are formulas again, and the workbook features are still there.
# Run scripts/write-scenario-outputs.mjs first. Usage: powershell -File excel-check.ps1
$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot '..\test-output' | Resolve-Path
$failures = New-Object System.Collections.Generic.List[string]
function Fail($msg) { $failures.Add($msg); Write-Output "  FAIL $msg" }

function Snapshot($wb) {
  $s = [ordered]@{}
  foreach ($ws in $wb.Worksheets) {
    $errors = @()
    $used = $ws.UsedRange
    $vals = $used.Value2
    if ($vals -is [Array]) {
      $r0 = $used.Row; $c0 = $used.Column
      for ($i = 1; $i -le $vals.GetLength(0); $i++) {
        for ($j = 1; $j -le $vals.GetLength(1); $j++) {
          if ($vals[$i, $j] -is [int]) { $errors += "R$($r0 + $i - 1)C$($c0 + $j - 1)" }
        }
      }
    }
    $validations = 0
    try { $validations = $ws.Cells.SpecialCells(-4174).Count } catch { }
    $s[$ws.Name] = [ordered]@{
      errors      = $errors
      validations = $validations
      formats     = $ws.Cells.FormatConditions.Count
      charts      = $ws.ChartObjects().Count
      visible     = $ws.Visible
      protected   = $ws.ProtectContents
    }
  }
  return $s
}

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  foreach ($name in 'ffos-template', 'ffos-template-excel') {
    Write-Output "== $name"
    $base = $null
    foreach ($pass in 'base', 'pass1', 'pass2') {
      $file = Join-Path $dir "$name.$pass.xlsx"
      try { $wb = $xl.Workbooks.Open($file, 0, $true) } catch { Fail "$name.$pass : Excel could not open it ($($_.Exception.Message))"; continue }
      if ($wb.Name -match 'Repaired') { Fail "$name.$pass : Excel repaired it" }
      $xl.CalculateFull()
      $snap = Snapshot $wb
      $me = $wb.Worksheets.Item('Month-End'); $gv = $wb.Worksheets.Item('Giving')
      if ($pass -eq 'base') {
        $base = $snap
        $n = ($snap.Values | ForEach-Object { $_.errors.Count } | Measure-Object -Sum).Sum
        Write-Output "  base opens; error cells (from NA() etc.): $n"
      } else {
        foreach ($sheet in $base.Keys) {
          $b = $base[$sheet]; $a = $snap[$sheet]
          $new = $a.errors | Where-Object { $b.errors -notcontains $_ }
          if ($new) { Fail "$name.$pass ${sheet}: new error cells $($new -join ',')" }
          foreach ($k in 'validations', 'formats', 'charts', 'visible', 'protected') {
            if ($a[$k] -ne $b[$k]) { Fail "$name.$pass ${sheet}: $k $($b[$k]) -> $($a[$k])" }
          }
        }
        $expect = @{ 'D7' = 1234.56; 'D8' = 250; 'D22' = 3000; 'D35' = 400; 'D38' = 10000; 'D45' = 500; 'D46' = -50; 'FX20' = 1 }
        if ($pass -eq 'pass1') { $expect['C7'] = 1000; $expect['D9'] = 75.1; $expect['D36'] = 150.25 }
        foreach ($ref in $expect.Keys) {
          $c = $me.Range($ref)
          if ($c.HasFormula -or $c.Value2 -ne $expect[$ref]) { Fail "$name.$pass Month-End!$ref = $($c.Value2) (formula: $($c.HasFormula)), expected $($expect[$ref])" }
        }
        if ($me.Range('D60').Value2 -ne 'Synthetic note: rent & <utilities> "test"') { Fail "$name.$pass note reads '$($me.Range('D60').Value2)'" }
        if ($gv.Range('A7').Text -eq '' -or $gv.Range('A7').Value2 -ne 46117 -or $gv.Range('C7').Value2 -ne 'Sample Charity A') { Fail "$name.$pass Giving row 7 wrong" }
        if ($pass -eq 'pass2') {
          foreach ($ref in 'C7', 'D9') { if (-not $me.Range($ref).HasFormula) { Fail "$name.pass2 Month-End!$ref is not a formula again" } }
          if ($me.Range('D36').Value2 -ne $null) { Fail "$name.pass2 D36 not blank" }
          if ($gv.Range('C8').Value2 -ne $null) { Fail "$name.pass2 Giving row 8 not cleared" }
        }
        Write-Output "  $pass checked; Giving A7 shows '$($gv.Range('A7').Text)'; Month-End D5 status '$($me.Range('D5').Text)'"
      }
      $wb.Close($false)
    }
  }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($failures.Count) { Write-Output "FAILED: $($failures.Count)"; exit 1 } else { Write-Output 'ALL EXCEL CHECKS PASSED' }
