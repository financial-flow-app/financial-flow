# Opens each scenario workbook in a private, hidden Excel, recalculates, and saves what Excel
# computed (Calc per month, metrics, accounts, goals, health checks, tithe quarters, Dashboard) as
# test/parity/<name>.excel.json. The Node parity test compares the app's engine to these files.
# Run scripts/parity-build.mjs first. Usage: powershell -File excel-parity.ps1
$ErrorActionPreference = 'Stop'
$inDir = Join-Path $PSScriptRoot '..\test-output\parity' | Resolve-Path
$outDir = Join-Path $PSScriptRoot '..\test\parity'
New-Item -ItemType Directory -Force $outDir | Out-Null
$errors = @{ -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'; -2146826265 = '#REF!'; -2146826259 = '#NAME?'; -2146826252 = '#NUM!'; -2146826246 = '#N/A' }
$inv = [Globalization.CultureInfo]::InvariantCulture

function J($v) {
  if ($null -eq $v) { return 'null' }
  if ($v -is [double]) { return $v.ToString('R', $inv) }
  if ($v -is [int]) { $e = $errors[$v]; if (-not $e) { $e = "#ERR$v" }; return '"' + $e + '"' }
  if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
  $s = [string]$v
  $s = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
  return '"' + $s + '"'
}
function Grid($range) {
  $v = $range.Value2
  $rows = @()
  for ($i = 1; $i -le $v.GetLength(0); $i++) {
    $cells = @()
    for ($j = 1; $j -le $v.GetLength(1); $j++) { $cells += (J $v[$i, $j]) }
    $rows += '[' + ($cells -join ',') + ']'
  }
  return '[' + ($rows -join ",`n") + ']'
}

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  foreach ($f in Get-ChildItem $inDir -Filter *.xlsx) {
    $wb = $xl.Workbooks.Open($f.FullName, 0, $true)
    $xl.CalculateFull()
    $calc = $wb.Worksheets.Item('Calc'); $hist = $wb.Worksheets.Item('History'); $dash = $wb.Worksheets.Item('Dashboard')
    $today = [double]$xl.Evaluate('TODAY()')
    $json = "{`n" +
      '"today": ' + (J $today) + ",`n" +
      '"calcHeaders": ' + (Grid $calc.Range('A5:FL5')) + ",`n" +
      '"calc": ' + (Grid $calc.Range('A6:FL183')) + ",`n" +
      '"metrics": ' + (Grid $calc.Range('EA5:EB38')) + ",`n" +
      '"taxMetrics": ' + (Grid $calc.Range('EA62:EB64')) + ",`n" +
      '"owedMetrics": ' + (Grid $calc.Range('EA65:EB67')) + ",`n" +
      '"accounts": ' + (Grid $calc.Range('EA40:EM59')) + ",`n" +
      '"goals": ' + (Grid $calc.Range('EA70:EH89')) + ",`n" +
      '"health": ' + (Grid $calc.Range('EA95:EA103')) + ",`n" +
      '"history": ' + (Grid $hist.Range('A25:G84')) + ",`n" +
      '"dashboard": ' + (Grid $dash.Range('A1:L49')) + "`n}`n"
    $name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    [IO.File]::WriteAllText((Join-Path $outDir "$name.excel.json"), $json, (New-Object Text.UTF8Encoding($false)))
    $wb.Close($false)
    Write-Output "saved test/parity/$name.excel.json"
  }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
