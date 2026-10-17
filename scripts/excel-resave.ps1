# Opens a workbook in a private, hidden Excel instance and saves a copy as .xlsx.
# Used to make the "saved by Excel" test fixture from the synthetic template.
# Usage: powershell -File excel-resave.ps1 <input.xlsx> <output.xlsx>
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$In = (Resolve-Path $In).Path
$Out = [IO.Path]::GetFullPath($Out)
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $wb = $xl.Workbooks.Open($In, 0, $true)
  $xl.CalculateFull()
  if (Test-Path $Out) { Remove-Item $Out }
  $wb.SaveAs($Out, 51)
  $wb.Close($false)
  Write-Output "saved $Out"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
