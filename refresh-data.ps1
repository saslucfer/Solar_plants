param(
    [string]$SourceFolder = $PSScriptRoot,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

# The first sheet of Natanz MDL.xlsx is a Shahreza list. Shahreza has its own file.
$plants = @(
    @{ File = 'Ali Abad MDL.xlsx'; Id = 'ali-abad'; Name = 'علی‌آباد'; Sheet = 1 },
    @{ File = 'Damghan MDL.xlsx'; Id = 'damghan'; Name = 'دامغان'; Sheet = 1 },
    @{ File = 'Evaz MDL.xlsx'; Id = 'evaz'; Name = 'اوز'; Sheet = 1 },
    @{ File = 'Hosnije MDL.xlsx'; Id = 'hosnije'; Name = 'حسنیجه'; Sheet = 1 },
    @{ File = 'Lar MDL.xlsx'; Id = 'lar'; Name = 'لار'; Sheet = 1 },
    @{ File = 'Marave Tape MDL.xlsx'; Id = 'marave-tape'; Name = 'مراوه‌تپه'; Sheet = 1 },
    @{ File = 'Natanz MDL.xlsx'; Id = 'natanz'; Name = 'نطنز'; Sheet = 2 },
    @{ File = 'Safi Abad MDL.xlsx'; Id = 'safi-abad'; Name = 'صفی‌آباد'; Sheet = 1 },
    @{ File = 'Shahreza MDL.xlsx'; Id = 'shahreza'; Name = 'شهرضا'; Sheet = 1 }
)

function Get-CellText($sheet, [int]$row, [int]$column) {
    return ([string]$sheet.Cells.Item($row, $column).Text).Trim()
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$records = [System.Collections.Generic.List[object]]::new()
$plantInfo = [System.Collections.Generic.List[object]]::new()
$seen = @{}

try {
    foreach ($plant in $plants) {
        $path = Join-Path $SourceFolder $plant.File
        if (-not (Test-Path -LiteralPath $path)) { throw "Missing workbook: $path" }
        $book = $excel.Workbooks.Open($path, 0, $true)
        try {
            $sheet = $book.Worksheets.Item($plant.Sheet)
            $used = $sheet.UsedRange
            $lastRow = $used.Row + $used.Rows.Count - 1
            $statusStart = if ((Get-CellText $sheet 1 3) -eq 'Document Number') { 4 } else { 3 }
            $category = 'سایر مدارک'
            $count = 0
            $unmarked = 0

            for ($row = 1; $row -le $lastRow; $row++) {
                $index = Get-CellText $sheet $row 1
                $title = Get-CellText $sheet $row 2
                $code = if ($statusStart -eq 4) { Get-CellText $sheet $row 3 } else { '' }
                $lastChange = Get-CellText $sheet $row ($statusStart + 4)
                $marks = @()
                for ($i = 0; $i -lt 4; $i++) {
                    if ((Get-CellText $sheet $row ($statusStart + $i)) -eq '*') { $marks += $i }
                }

                if ($index -and -not $title -and $index -notmatch '^(Row|\d+)$' -and $row -gt 2) {
                    $category = $index
                    continue
                }
                if (-not $title) { continue }
                if ($title -match '^(Civil|Electrical|General|Structure)$' -and -not $code -and $marks.Count -eq 0) {
                    $category = $title
                    continue
                }
                if ($title -match '^Master Document Registration' -and $marks.Count -eq 0 -and $row -le 3) { continue }
                if ($row -le 2 -and $marks.Count -eq 0) { continue }

                $status = if ($marks.Count -eq 1) { @('approved', 'pending', 'not-sent', 'revision')[$marks[0]] } else { 'unknown' }
                if ($status -eq 'unknown') { $unmarked++ }
                $rawKey = ($plant.Id + '|' + $code.ToLowerInvariant() + '|' + $title.ToLowerInvariant())
                if (-not $seen.ContainsKey($rawKey)) { $seen[$rawKey] = 0 }
                $seen[$rawKey]++
                $id = $rawKey + '|' + $seen[$rawKey]

                $records.Add([ordered]@{
                    id = $id
                    plantId = $plant.Id
                    category = $category
                    title = $title
                    code = $code
                    status = $status
                    lastChange = $lastChange
                    sourceRow = $row
                })
                $count++
            }
            $plantInfo.Add([ordered]@{
                id = $plant.Id
                name = $plant.Name
                file = $plant.File
                sheet = [string]$sheet.Name
                count = $count
                unmarked = $unmarked
            })
        }
        finally {
            $book.Close($false)
        }
    }
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

$payload = [ordered]@{
    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
    plants = @($plantInfo.ToArray())
    documents = @($records.ToArray())
}
if ($PassThru) {
    return $payload
}
$json = ConvertTo-Json -InputObject $payload -Depth 7 -Compress
$output = Join-Path $PSScriptRoot 'data.js'
$utf8 = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($output, 'window.MDL_DATA = ' + $json + ';', $utf8)
Write-Output "Created $output with $($records.Count) documents across $($plantInfo.Count) plants."
