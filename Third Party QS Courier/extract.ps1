$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Read-Workbook($path, $label) {
    Write-Host "=== $label ==="
    $wb = $excel.Workbooks.Open($path)
    foreach ($sheet in $wb.Sheets) {
        Write-Host "--- Sheet: $($sheet.Name) ---"
        $usedRange = $sheet.UsedRange
        $rows = $usedRange.Rows.Count
        $cols = $usedRange.Columns.Count
        for ($r = 1; $r -le [Math]::Min($rows, 100); $r++) {
            $line = ""
            for ($c = 1; $c -le $cols; $c++) {
                $val = $sheet.Cells($r, $c).Text
                $line += "[$val]"
            }
            Write-Host $line
        }
    }
    $wb.Close($false)
}

Read-Workbook "C:\Users\YUSRI-VICTUS\Desktop\rate-shipment-calculator\QS Smart Rate.xlsx" "QS Smart Rate"
Read-Workbook "C:\Users\YUSRI-VICTUS\Desktop\rate-shipment-calculator\2026 Quantium Solutions war surcharge (32%).xlsx" "War Surcharge"

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
