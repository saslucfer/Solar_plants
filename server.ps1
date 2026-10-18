param(
    [string]$SourceFolder = $PSScriptRoot,
    [int]$Port = 18765,
    [switch]$OpenBrowser
)

$ErrorActionPreference = 'Stop'
$SourceFolder = [IO.Path]::GetFullPath($SourceFolder)
$utf8 = [Text.UTF8Encoding]::new($false)
$script:currentData = $null
$script:currentVersion = ''
$statusKeys = @('approved', 'pending', 'not-sent', 'revision')
$allowedStatuses = $statusKeys + @('unknown')

function Get-WorkbookVersion {
    $names = @('Ali Abad MDL.xlsx', 'Damghan MDL.xlsx', 'Evaz MDL.xlsx', 'Hosnije MDL.xlsx', 'Lar MDL.xlsx', 'Marave Tape MDL.xlsx', 'Natanz MDL.xlsx', 'Safi Abad MDL.xlsx', 'Shahreza MDL.xlsx')
    $parts = foreach ($name in $names) {
        $file = Get-Item -LiteralPath (Join-Path $SourceFolder $name)
        "$name/$($file.Length)/$($file.LastWriteTimeUtc.Ticks)"
    }
    $bytes = $utf8.GetBytes(($parts -join '|'))
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $hash.Dispose() }
}

function Update-Cache {
    $version = Get-WorkbookVersion
    if ($null -eq $script:currentData -or $version -ne $script:currentVersion) {
        $script:currentData = & (Join-Path $PSScriptRoot 'refresh-data.ps1') -SourceFolder $SourceFolder -PassThru
        $script:currentVersion = Get-WorkbookVersion
    }
}

function Send-Bytes($response, [int]$status, [byte[]]$bytes, [string]$contentType) {
    $response.StatusCode = $status
    $response.ContentType = $contentType
    $response.ContentLength64 = $bytes.Length
    $response.Headers.Add('Cache-Control', 'no-store')
    try { $response.OutputStream.Write($bytes, 0, $bytes.Length) }
    finally { $response.Close() }
}

function Send-Json($response, [int]$status, $body) {
    $json = ConvertTo-Json -InputObject $body -Depth 8 -Compress
    Send-Bytes $response $status $utf8.GetBytes($json) 'application/json; charset=utf-8'
}

function Get-CellText($sheet, [int]$row, [int]$column) {
    return ([string]$sheet.Cells.Item($row, $column).Text).Trim()
}

function Convert-MdlDate([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    $digits = '۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩'
    for ($i = 0; $i -lt 10; $i++) {
        $value = $value.Replace([string]$digits[$i], [string]$i).Replace([string]$digits[$i + 10], [string]$i)
    }
    $value = $value.Trim().Replace('-', '/').Replace('.', '/')
    if ($value -notmatch '^(\d{4})/(\d{1,2})/(\d{1,2})$') { throw '[400] تاریخ شمسی باید به صورت ۱۴۰۵/۰۷/۲۶ باشد.' }
    $year = [int]$Matches[1]; $month = [int]$Matches[2]; $day = [int]$Matches[3]
    if ($year -lt 1200 -or $year -gt 1600) { throw '[400] سال تاریخ شمسی معتبر نیست.' }
    try { [void]([Globalization.PersianCalendar]::new().ToDateTime($year, $month, $day, 0, 0, 0, 0)) }
    catch { throw '[400] تاریخ شمسی معتبر نیست.' }
    return ('{0:D4}/{1:D2}/{2:D2}' -f $year, $month, $day)
}

function Save-Document($request) {
    if (-not $request -or -not $request.id -or $allowedStatuses -notcontains [string]$request.status) { throw '[400] شناسه یا وضعیت مدرک معتبر نیست.' }
    if ($request.PSObject.Properties.Name -notcontains 'lastChange' -or $request.PSObject.Properties.Name -notcontains 'expectedStatus' -or $request.PSObject.Properties.Name -notcontains 'expectedLastChange') { throw '[400] داده‌های تغییر ناقص است.' }
    $newDate = Convert-MdlDate ([string]$request.lastChange)
    Update-Cache
    $doc = $script:currentData.documents | Where-Object { $_.id -eq [string]$request.id } | Select-Object -First 1
    if (-not $doc) { throw '[404] مدرک پیدا نشد؛ صفحه را تازه‌سازی کنید.' }
    $plant = $script:currentData.plants | Where-Object { $_.id -eq $doc.plantId } | Select-Object -First 1
    $path = Join-Path $SourceFolder $plant.file
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $book = $null
    try {
        $book = $excel.Workbooks.Open($path, 0, $false)
        if ($book.ReadOnly) { throw '[423] فایل اکسل برای ویرایش باز است. آن را ذخیره و ببندید، سپس دوباره تلاش کنید.' }
        $sheet = $book.Worksheets.Item($plant.sheet)
        $row = [int]$doc.sourceRow
        $statusStart = if ((Get-CellText $sheet 1 3) -eq 'Document Number') { 4 } else { 3 }
        $title = Get-CellText $sheet $row 2
        $code = if ($statusStart -eq 4) { Get-CellText $sheet $row 3 } else { '' }
        if ($title -ne $doc.title -or $code -ne $doc.code) { throw '[409] ردیف مدرک در اکسل جابه‌جا شده است. صفحه را تازه‌سازی کنید.' }
        $marks = @()
        for ($i = 0; $i -lt 4; $i++) {
            if ((Get-CellText $sheet $row ($statusStart + $i)) -eq '*') { $marks += $i }
        }
        $actualStatus = if ($marks.Count -eq 1) { $statusKeys[$marks[0]] } else { 'unknown' }
        $actualDate = Get-CellText $sheet $row ($statusStart + 4)
        if ($actualStatus -ne [string]$request.expectedStatus -or $actualDate -ne [string]$request.expectedLastChange) { throw '[409] این مدرک در اکسل تغییر کرده است. صفحه را تازه‌سازی کنید.' }
        for ($i = 0; $i -lt 4; $i++) {
            $sheet.Cells.Item($row, $statusStart + $i).Value2 = if ($statusKeys[$i] -eq [string]$request.status) { '*' } else { '' }
        }
        $dateColumn = $statusStart + 4
        if (-not (Get-CellText $sheet 1 $dateColumn)) {
            $sheet.Cells.Item(1, $dateColumn).Value2 = 'تاریخ آخرین تغییر وضعیت (شمسی)'
            $sheet.Columns.Item($dateColumn).ColumnWidth = 25
        }
        $sheet.Cells.Item($row, $dateColumn).NumberFormat = '@'
        $sheet.Cells.Item($row, $dateColumn).Value2 = $newDate
        $book.Save()
        $doc.status = [string]$request.status
        $doc.lastChange = $newDate
        $script:currentData.generatedAt = (Get-Date).ToUniversalTime().ToString('o')
        $script:currentVersion = Get-WorkbookVersion
        return [ordered]@{ document = $doc; version = $script:currentVersion }
    }
    finally {
        if ($null -ne $book) { $book.Close($false) }
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}

$prefix = "http://localhost:$Port/"
$listener = [Net.HttpListener]::new()
$listener.Prefixes.Add($prefix)
try {
    $listener.Start()
    Write-Output "Dashboard running at $prefix"
    Write-Output 'Keep this PowerShell window open. Close it to stop the dashboard.'
    if ($OpenBrowser) { Start-Process $prefix }
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        try {
            $path = $request.Url.AbsolutePath
            if ($request.HttpMethod -eq 'GET' -and $path -eq '/api/version') {
                Send-Json $response 200 ([ordered]@{ version = Get-WorkbookVersion })
            }
            elseif ($request.HttpMethod -eq 'GET' -and $path -eq '/api/data') {
                Update-Cache
                Send-Json $response 200 ([ordered]@{ version = $script:currentVersion; data = $script:currentData })
            }
            elseif ($request.HttpMethod -eq 'POST' -and $path -eq '/api/document') {
                $origin = [string]$request.Headers['Origin']
                if ($origin -and $origin -ne "http://localhost:$Port") { throw '[403] درخواست از مبدأ مجاز نیست.' }
                if ($request.ContentType -notlike 'application/json*' -or $request.ContentLength64 -gt 10000) { throw '[400] قالب درخواست معتبر نیست.' }
                $reader = [IO.StreamReader]::new($request.InputStream, $utf8)
                try { $body = $reader.ReadToEnd() | ConvertFrom-Json }
                finally { $reader.Dispose() }
                $result = Save-Document $body
                Send-Json $response 200 $result
            }
            elseif ($request.HttpMethod -eq 'GET') {
                $filename = switch ($path) { '/' { 'index.html' } '/index.html' { 'index.html' } '/style.css' { 'style.css' } '/app.js' { 'app.js' } '/data.js' { 'data.js' } default { $null } }
                if (-not $filename) { Send-Json $response 404 ([ordered]@{ error = 'فایل پیدا نشد.' }); continue }
                $contentType = if ($filename -eq 'index.html') { 'text/html; charset=utf-8' } elseif ($filename -eq 'style.css') { 'text/css; charset=utf-8' } else { 'application/javascript; charset=utf-8' }
                Send-Bytes $response 200 ([IO.File]::ReadAllBytes((Join-Path $PSScriptRoot $filename))) $contentType
            }
            else { Send-Json $response 405 ([ordered]@{ error = 'روش درخواست مجاز نیست.' }) }
        }
        catch {
            $message = $_.Exception.Message
            $status = 500
            if ($message -match '^\[(\d{3})\]\s*(.*)$') { $status = [int]$Matches[1]; $message = $Matches[2] }
            Write-Warning $message
            try { Send-Json $response $status ([ordered]@{ error = $message }) } catch { }
        }
    }
}
finally {
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
}
