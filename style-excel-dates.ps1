param([string]$SourceFolder = $PSScriptRoot)

$ErrorActionPreference = 'Stop'
$SourceFolder = [IO.Path]::GetFullPath($SourceFolder)
$files = @('Ali Abad MDL.xlsx', 'Damghan MDL.xlsx', 'Evaz MDL.xlsx', 'Hosnije MDL.xlsx', 'Lar MDL.xlsx', 'Marave Tape MDL.xlsx', 'Natanz MDL.xlsx', 'Safi Abad MDL.xlsx', 'Shahreza MDL.xlsx')

function Get-OleColor([string]$hex) {
    $red = [Convert]::ToInt32($hex.Substring(1, 2), 16)
    $green = [Convert]::ToInt32($hex.Substring(3, 2), 16)
    $blue = [Convert]::ToInt32($hex.Substring(5, 2), 16)
    return ($red + ($green -shl 8) + ($blue -shl 16))
}

foreach ($name in $files) {
    if (-not (Test-Path -LiteralPath (Join-Path $SourceFolder $name))) { throw "فایل اکسل پیدا نشد: $name" }
}

$backupFolder = Join-Path $SourceFolder ('workbook-backups\before-date-style-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $backupFolder -ErrorAction Stop | Out-Null
foreach ($name in $files) {
    Copy-Item -LiteralPath (Join-Path $SourceFolder $name) -Destination (Join-Path $backupFolder $name) -ErrorAction Stop
}

$navy = Get-OleColor '#102E38'
$mint = Get-OleColor '#D5F2E9'
$paleMint = Get-OleColor '#F1F9F5'
$teal = Get-OleColor '#1D9278'
$ink = Get-OleColor '#17333A'
$line = Get-OleColor '#DCEBE6'
$white = Get-OleColor '#FFFFFF'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$styled = 0
try {
    foreach ($name in $files) {
        $workbook = $excel.Workbooks.Open((Join-Path $SourceFolder $name), 0, $false)
        try {
            if ($workbook.ReadOnly) { throw "فایل برای ویرایش در Excel باز است: $name" }
            foreach ($sheet in $workbook.Worksheets) {
                $dateColumn = if (([string]$sheet.Cells.Item(1, 3).Text).Trim() -eq 'Document Number') { 8 } else { 7 }
                $lastRow = $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1
                $sheet.Columns.Item($dateColumn).ColumnWidth = 30

                $header = $sheet.Cells.Item(1, $dateColumn)
                $header.Value2 = 'آخرین تغییر وضعیت'
                $header.Interior.Color = $navy
                $header.Font.Name = 'Tahoma'
                $header.Font.Size = 12
                $header.Font.Bold = $true
                $header.Font.Color = $white
                $header.HorizontalAlignment = -4108
                $header.VerticalAlignment = -4108
                $header.WrapText = $true
                $header.Borders.Item(9).LineStyle = 1
                $header.Borders.Item(9).Color = $teal

                $hint = $sheet.Cells.Item(2, $dateColumn)
                $hint.Value2 = 'تاریخ شمسی: ۱۴۰۵/۰۷/۲۶'
                $hint.Interior.Color = $mint
                $hint.Font.Name = 'Tahoma'
                $hint.Font.Size = 9
                $hint.Font.Bold = $false
                $hint.Font.Color = $ink
                $hint.HorizontalAlignment = -4108
                $hint.VerticalAlignment = -4108
                $hint.WrapText = $true

                if ($sheet.Rows.Item(1).RowHeight -lt 38) { $sheet.Rows.Item(1).RowHeight = 38 }
                if ($sheet.Rows.Item(2).RowHeight -lt 26) { $sheet.Rows.Item(2).RowHeight = 26 }

                if ($lastRow -ge 3) {
                    $body = $sheet.Range($sheet.Cells.Item(3, $dateColumn), $sheet.Cells.Item($lastRow, $dateColumn))
                    $body.NumberFormat = '@'
                    $body.Interior.Color = $paleMint
                    $body.Font.Name = 'Tahoma'
                    $body.Font.Size = 10
                    $body.Font.Color = $ink
                    $body.HorizontalAlignment = -4108
                    $body.VerticalAlignment = -4108
                    foreach ($edge in @(7, 10, 12)) {
                        $body.Borders.Item($edge).LineStyle = 1
                        $body.Borders.Item($edge).Color = $line
                    }
                }
                $styled++
            }
            $workbook.Save()
        }
        finally { $workbook.Close($false) }
    }
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

Write-Output "Styled $styled MDL sheets in $($files.Count) workbooks."
Write-Output "Backups: $backupFolder"
