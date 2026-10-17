# Checks that Excel opens a stamped export without repair, sees the stamp as a custom document
# property, and keeps it when it saves the file. Writes test-output/stamp-saved.xlsx for Node to read.
# Usage: node scripts/write-stamp-sample.mjs; powershell -File scripts/excel-stamp-check.ps1
$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot '..\test-output' | Resolve-Path
$in = Join-Path $dir 'stamp.xlsx'
$out = Join-Path $dir 'stamp-saved.xlsx'
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  $wb = $xl.Workbooks.Open($in, 0, $false)
  if ($wb.Name -match 'Repaired') { throw 'Excel repaired the file' }
  # PowerShell can't use Office's property collection directly; call it late-bound by name.
  $get = [Reflection.BindingFlags]::GetProperty
  $props = [System.__ComObject].InvokeMember('CustomDocumentProperties', $get, $null, $wb, $null)
  $value = $null
  try {
    $p = [System.__ComObject].InvokeMember('Item', $get, $null, $props, @('FinancialFlowExport'))
    $value = [System.__ComObject].InvokeMember('Value', $get, $null, $p, $null)
  } catch { }
  if (-not $value) { throw 'Excel does not see the stamp' }
  Write-Output "Excel sees FinancialFlowExport = $value"
  $wb.Worksheets.Item('Month-End').Range('E7').Value2 = 1  # an edit, as if made in Excel
  if (Test-Path $out) { Remove-Item $out }
  $wb.SaveAs($out, 51)
  $wb.Close($false)
  Write-Output 'saved after an edit'
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
