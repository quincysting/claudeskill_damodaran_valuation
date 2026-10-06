$ErrorActionPreference = 'Stop'
$path = Join-Path $PSScriptRoot 'SPCX_Valuation_2026-09-25.xlsx'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($path)
$m = [System.Reflection.Missing]::Value
$s = $wb.Worksheets.Item('Scenarios');   $null = $s.Range('N7:O19').Table($m, $s.Range('D5'))
$t = $wb.Worksheets.Item('Sensitivity')
$null = $t.Range('B13:G22').Table($t.Range('C6'), $t.Range('C5'))
$null = $t.Range('B27:G34').Table($t.Range('C8'), $t.Range('C7'))
$null = $t.Range('B39:D46').Table($m, $t.Range('C9'))
foreach ($r in @($t.Range('C14:G22'), $t.Range('C28:G34'), $t.Range('C40:C46'), $s.Range('O8:O19'))) {
  $r.Font.Name = 'Arial'; $r.Font.Size = 10; $r.NumberFormat = '$#,##0.00_);($#,##0.00)' }
$t.Range('D40:D46').Font.Name = 'Arial'; $t.Range('D40:D46').Font.Size = 10; $t.Range('D40:D46').NumberFormat = '0.00%'
$xl.CalculateFull()
$wb.Worksheets.Item('Cover').Activate()
$wb.Save(); $wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Output 'finalized'
