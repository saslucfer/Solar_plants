function ConvertTo-ReportHtmlText($value) {
    return [Net.WebUtility]::HtmlEncode([string]$value)
}

function ConvertTo-ReportDay([string]$value) {
    $digits = '۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩'
    for ($i = 0; $i -lt 10; $i++) { $value = $value.Replace([string]$digits[$i], [string]$i).Replace([string]$digits[$i + 10], [string]$i) }
    $value = $value.Trim().Replace('-', '/').Replace('.', '/')
    if ($value -notmatch '^(\d{4})/(\d{1,2})/(\d{1,2})$') { return $null }
    if ([int]$Matches[1] -lt 1200 -or [int]$Matches[1] -gt 1600) { return $null }
    try { return [Globalization.PersianCalendar]::new().ToDateTime([int]$Matches[1], [int]$Matches[2], [int]$Matches[3], 0, 0, 0, 0).Date }
    catch { return $null }
}

function Format-ReportTimestamp([string]$value) {
    try {
        $date = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::Parse($value).ToUniversalTime(), [TimeZoneInfo]::FindSystemTimeZoneById('Iran Standard Time'))
        $calendar = [Globalization.PersianCalendar]::new()
        return Convert-ToPersianDigits ('{0:D4}/{1:D2}/{2:D2} - {3:HH:mm}' -f $calendar.GetYear($date), $calendar.GetMonth($date), $calendar.GetDayOfMonth($date), $date)
    } catch { return $value }
}

function Add-ReportDocumentTable($builder, $documents) {
    [void]$builder.Append('<table class="documents"><thead><tr><th>ردیف اکسل</th><th>نام مدرک</th><th>شماره مدرک</th><th>بخش</th><th>وضعیت</th><th>آخرین تغییر وضعیت</th></tr></thead><tbody>')
    foreach ($doc in $documents) {
        $date = if ([string]::IsNullOrWhiteSpace([string]$doc.lastChange)) { 'ثبت نشده' } else { Convert-ToPersianDigits ([string]$doc.lastChange) }
        $code = if ($doc.code) { $doc.code } else { '—' }
        $status = if ($script:reportLabels.ContainsKey([string]$doc.status)) { $script:reportLabels[[string]$doc.status] } else { 'نامشخص' }
        [void]$builder.Append('<tr class="document-row"><td>' + (Convert-ToPersianDigits ([string]$doc.sourceRow)) + '</td><td dir="auto" class="document-title">' + (ConvertTo-ReportHtmlText $doc.title) + '</td><td class="code" dir="ltr">' + (ConvertTo-ReportHtmlText $code) + '</td><td dir="auto">' + (ConvertTo-ReportHtmlText $doc.category) + '</td><td class="status-' + (ConvertTo-ReportHtmlText $doc.status) + '">' + $status + '</td><td>' + (ConvertTo-ReportHtmlText $date) + '</td></tr>')
    }
    [void]$builder.Append('</tbody></table>')
}

function New-FullReportHtml($data, [object[]]$notes, [DateTime]$generatedUtc = [DateTime]::UtcNow) {
    $script:reportLabels = @{ approved = 'تأیید شده'; pending = 'در انتظار تأیید'; 'not-sent' = 'ارسال نشده'; revision = 'در انتظار اصلاح'; unknown = 'نامشخص' }
    $tehran = [TimeZoneInfo]::ConvertTimeFromUtc($generatedUtc, [TimeZoneInfo]::FindSystemTimeZoneById('Iran Standard Time'))
    $cutoff = $tehran.Date.AddDays(-7)
    $overdue = @{}
    $missing = 0
    foreach ($doc in $data.documents) {
        $day = ConvertTo-ReportDay ([string]$doc.lastChange)
        if ($null -eq $day) { $missing++ }
        elseif ($doc.status -ne 'approved' -and $day -lt $cutoff) { $overdue[[string]$doc.id] = $true }
    }
    $builder = [Text.StringBuilder]::new()
    [void]$builder.Append(@'
<!doctype html><html lang="fa" dir="rtl"><head><meta charset="utf-8"><title>گزارش جامع مدارک نیروگاه‌های خورشیدی</title><style>
@font-face{font-family:ReportPersian;src:url('b-zar.ttf') format('truetype');font-weight:400;unicode-range:U+0600-06FF,U+0750-077F,U+08A0-08FF,U+200C-200F,U+FB50-FDFF,U+FE70-FEFF}
@font-face{font-family:ReportPersian;src:url('b-zar-bold.ttf') format('truetype');font-weight:700;unicode-range:U+0600-06FF,U+0750-077F,U+08A0-08FF,U+200C-200F,U+FB50-FDFF,U+FE70-FEFF}
@page{size:A4 landscape;margin:12mm 11mm 16mm;@bottom-center{content:"صفحه " counter(page,persian) " از " counter(pages,persian);direction:rtl;font:10pt ReportPersian,Arial;color:#667b80}}
*{box-sizing:border-box}body{margin:0;color:#193d46;background:white;font:13pt/1.6 ReportPersian,"Segoe UI",Tahoma,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
h1{font-size:28pt;margin:0 0 5mm}h2{font-size:22pt;margin:0 0 4mm}h3{font-size:17pt;margin:5mm 0 2mm;break-after:avoid}p{margin:2mm 0}header{border-bottom:3px solid #298673;padding-bottom:5mm;margin-bottom:5mm}.meta{color:#5c747b;font-size:11pt}.summary{display:flex;gap:4mm;margin:5mm 0}.metric{flex:1;padding:3mm;border:1px solid #dce7e3;background:#f3f8f5;border-radius:3mm}.metric strong{display:block;font-size:23pt}.metric span{font-size:12pt}.project{break-before:page}.project-head{padding:4mm;background:#edf5f1;border-right:4px solid #298673;margin-bottom:4mm}.project-head h2{margin:0}.empty{color:#6f8185;padding:3mm;background:#f8faf9;border:1px solid #e5ece8}.note{border-bottom:1px solid #dce6e2;padding:3mm 0}.note h4{font-size:15pt;margin:1mm 0;break-after:avoid}.note-body{white-space:pre-wrap;overflow-wrap:anywhere;unicode-bidi:plaintext;font-size:13pt}.note-meta{color:#627e84;font-size:11pt;break-after:avoid}
table{width:100%;border-collapse:collapse;table-layout:fixed;margin:3mm 0 5mm;font-size:12pt}thead{display:table-header-group}th,td{border:1px solid #d9e4df;padding:2mm 2.3mm;vertical-align:top;overflow-wrap:anywhere}th{background:#eaf2ef;font-weight:700;color:#25534f}tr{break-inside:avoid}tbody tr:nth-child(even){background:#f8fbf9}.documents th:nth-child(1){width:6%}.documents th:nth-child(2){width:28%}.documents th:nth-child(3){width:22%}.documents th:nth-child(4){width:13%}.documents th:nth-child(5){width:16%}.documents th:nth-child(6){width:15%}.code{font:9pt/1.65 Consolas,"Courier New",monospace;text-align:left}.document-title{font-size:10pt;line-height:1.65}.status-approved{color:#187257}.status-pending{color:#976315}.status-revision{color:#a75445}.explanation{padding:3mm;border-right:3px solid #d2a553;background:#fff9ec;font-size:12pt}
body{font-size:12pt;line-height:1.4}h1{font-size:25pt;line-height:1.3;margin-bottom:2mm}h2{font-size:20pt;margin-bottom:2mm}h3{font-size:16pt;margin:4mm 0 2mm}header{padding-bottom:3mm;margin-bottom:3mm}.meta{font-size:10pt}.summary{margin:3mm 0}.metric{padding:2mm 3mm}.metric strong{font-size:21pt}.metric span{font-size:11pt}.explanation{padding:2mm;font-size:11pt}table{font-size:11.5pt;margin:2mm 0 4mm}th,td{padding:1.3mm 1.8mm}.code{font-size:8.5pt;line-height:1.45}.document-title{font-size:9.5pt;line-height:1.45}body>table{font-size:10.5pt}.project-head{padding:3mm}.note-body{font-size:12pt}
</style></head><body>
'@)
    [void]$builder.Append('<header><h1>گزارش جامع مدارک نیروگاه‌های خورشیدی</h1><p>اطلس مدارک — همهٔ پروژه‌ها و تمام داده‌های ذخیره‌شده</p><p class="meta">تاریخ تهیه گزارش (تهران): ' + (Format-ReportTimestamp $generatedUtc.ToString('o')) + '</p><p class="meta">آخرین خواندن مدارک از اکسل: ' + (ConvertTo-ReportHtmlText (Format-ReportTimestamp ([string]$data.generatedAt))) + '</p></header>')
    $approved = @($data.documents | Where-Object { $_.status -eq 'approved' }).Count
    $rate = if (@($data.documents).Count) { [Math]::Round(100 * $approved / @($data.documents).Count) } else { 0 }
    $approvedLabel = 'تأیید شده (' + (Convert-ToPersianDigits ([string]$rate)) + '٪)'
    [void]$builder.Append('<div class="summary">')
    foreach ($metric in @(@('نیروگاه', @($data.plants).Count), @('کل مدارک', @($data.documents).Count), @($approvedLabel, $approved), @('نیازمند پیگیری بیش از یک هفته', $overdue.Count), @('یادداشت، کامنت و گزارش', @($notes).Count))) {
        [void]$builder.Append('<div class="metric"><strong>' + (Convert-ToPersianDigits ([string]$metric[1])) + '</strong><span>' + $metric[0] + '</span></div>')
    }
    [void]$builder.Append('</div><p class="explanation">این گزارش شامل همهٔ نیروگاه‌ها و تمام ردیف‌های مدارک است و به فیلترها یا صفحهٔ انتخاب‌شده در داشبورد محدود نمی‌شود. ' + (Convert-ToPersianDigits ([string]$missing)) + ' مدرک بدون تاریخ معتبر است. در فهرست پیگیری، فقط مدارک تأییدنشده با تاریخ معتبرِ قدیمی‌تر از هفت روز پیشِ امروز تهران آمده‌اند. پیش‌نویس‌های ثبت‌نشده در گزارش نیستند.</p><h2>خلاصهٔ وضعیت پروژه‌ها</h2><table><thead><tr><th>نیروگاه</th><th>کل مدارک</th><th>تأیید شده</th><th>در انتظار تأیید</th><th>در انتظار اصلاح</th><th>ارسال نشده</th><th>نامشخص</th><th>نیازمند پیگیری</th><th>توضیحات ثبت‌شده</th></tr></thead><tbody>')
    foreach ($plant in $data.plants) {
        $docs = @($data.documents | Where-Object { $_.plantId -eq $plant.id })
        [void]$builder.Append('<tr><td>' + (ConvertTo-ReportHtmlText $plant.name) + '</td><td>' + (Convert-ToPersianDigits ([string]$docs.Count)) + '</td>')
        foreach ($status in @('approved', 'pending', 'revision', 'not-sent', 'unknown')) { [void]$builder.Append('<td>' + (Convert-ToPersianDigits ([string]@($docs | Where-Object { $_.status -eq $status }).Count)) + '</td>') }
        [void]$builder.Append('<td>' + (Convert-ToPersianDigits ([string]@($docs | Where-Object { $overdue.ContainsKey([string]$_.id) }).Count)) + '</td><td>' + (Convert-ToPersianDigits ([string]@($notes | Where-Object { $_.plantId -eq $plant.id }).Count)) + '</td></tr>')
    }
    [void]$builder.Append('</tbody></table>')
    $kinds = @{ note = 'یادداشت'; comment = 'کامنت'; report = 'گزارش' }
    foreach ($plant in $data.plants) {
        $docs = @($data.documents | Where-Object { $_.plantId -eq $plant.id } | Sort-Object { [int]$_.sourceRow })
        $plantNotes = @($notes | Where-Object { $_.plantId -eq $plant.id } | Sort-Object createdAt)
        $plantApproved = @($docs | Where-Object { $_.status -eq 'approved' }).Count
        $plantRate = if ($docs.Count) { [Math]::Round(100 * $plantApproved / $docs.Count) } else { 0 }
        [void]$builder.Append('<section class="project"><div class="project-head"><h2>نیروگاه ' + (ConvertTo-ReportHtmlText $plant.name) + '</h2><p class="meta">فایل منبع: <bdi dir="ltr">' + (ConvertTo-ReportHtmlText $plant.file) + '</bdi> | برگه: <bdi dir="auto">' + (ConvertTo-ReportHtmlText $plant.sheet) + '</bdi> | تعداد مدارک: ' + (Convert-ToPersianDigits ([string]$docs.Count)) + ' | تأیید شده: ' + (Convert-ToPersianDigits ([string]$plantRate)) + '٪</p></div><h3>فهرست کامل مدارک</h3>')
        Add-ReportDocumentTable $builder $docs
        [void]$builder.Append('<h3>مدارک تأییدنشده با گذشت بیش از یک هفته</h3>')
        foreach ($group in @('pending', 'revision', 'other')) {
            $items = @($docs | Where-Object { $overdue.ContainsKey([string]$_.id) -and (($group -eq 'other' -and $_.status -notin @('pending', 'revision')) -or $_.status -eq $group) })
            if ($group -eq 'other' -and -not $items.Count) { continue }
            $label = if ($group -eq 'other') { 'سایر وضعیت‌های تأییدنشده' } else { $script:reportLabels[$group] }
            [void]$builder.Append('<h3>' + $label + ' — ' + (Convert-ToPersianDigits ([string]$items.Count)) + ' مدرک</h3>')
            if ($items.Count) { Add-ReportDocumentTable $builder $items }
            else { [void]$builder.Append('<p class="empty">مدرک واجد شرایطی وجود ندارد.</p>') }
        }
        [void]$builder.Append('<h3>یادداشت‌ها، کامنت‌ها و گزارش‌های پروژه — ' + (Convert-ToPersianDigits ([string]$plantNotes.Count)) + ' مورد</h3>')
        if (-not $plantNotes.Count) { [void]$builder.Append('<p class="empty">برای این نیروگاه توضیحی ثبت نشده است.</p>') }
        foreach ($note in $plantNotes) {
            [void]$builder.Append('<article class="note"><p class="note-meta">نوع: ' + $kinds[[string]$note.kind] + ' | ثبت: ' + (ConvertTo-ReportHtmlText (Format-ReportTimestamp ([string]$note.createdAt))) + ' | آخرین ویرایش: ' + (ConvertTo-ReportHtmlText (Format-ReportTimestamp ([string]$note.updatedAt))) + '</p>')
            if ($note.title) { [void]$builder.Append('<h4 dir="auto">' + (ConvertTo-ReportHtmlText $note.title) + '</h4>') }
            [void]$builder.Append('<div class="note-body" dir="auto">' + (ConvertTo-ReportHtmlText $note.body) + '</div></article>')
        }
        [void]$builder.Append('</section>')
    }
    [void]$builder.Append('</body></html>')
    return $builder.ToString()
}

function New-FullReportPdf($data, [object[]]$notes) {
    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
    )
    $browser = $candidates | Where-Object { [IO.File]::Exists($_) } | Select-Object -First 1
    if (-not $browser) { throw '[503] برای ساخت گزارش PDF باید Microsoft Edge یا Google Chrome روی این رایانه نصب باشد.' }
    $folder = Join-Path ([IO.Path]::GetTempPath()) ('AtlasMdlReport-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($folder)
    $process = $null
    try {
        foreach ($font in @('b-zar.ttf', 'b-zar-bold.ttf')) { [IO.File]::Copy((Join-Path $PSScriptRoot "fonts\$font"), (Join-Path $folder $font)) }
        $htmlPath = Join-Path $folder 'report.html'
        $pdfPath = Join-Path $folder 'report.pdf'
        [IO.File]::WriteAllText($htmlPath, (New-FullReportHtml $data $notes), [Text.UTF8Encoding]::new($false))
        $info = [Diagnostics.ProcessStartInfo]::new()
        $info.FileName = $browser
        $info.Arguments = '--headless=new --disable-gpu --no-first-run --no-default-browser-check --disable-extensions --disable-background-networking --no-pdf-header-footer --timeout=20000 --user-data-dir="' + (Join-Path $folder 'profile') + '" --print-to-pdf="' + $pdfPath + '" "' + ([Uri]$htmlPath).AbsoluteUri + '"'
        $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
        $process = [Diagnostics.Process]::Start($info)
        $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(45000)) { $process.Kill(); throw '[504] ساخت گزارش PDF بیش از حد طول کشید؛ دوباره تلاش کنید.' }
        if (-not [IO.File]::Exists($pdfPath)) { throw '[500] فایل PDF ساخته نشد؛ مرورگر را به‌روز کنید و دوباره تلاش کنید.' }
        $bytes = [IO.File]::ReadAllBytes($pdfPath)
        if ($bytes.Length -lt 5 -or [Text.Encoding]::ASCII.GetString($bytes, 0, 5) -ne '%PDF-') { throw '[500] خروجی گزارش PDF معتبر نیست.' }
        return ,$bytes
    }
    finally {
        if ($process) { $process.Dispose() }
        # This is a unique directory created above, never a user-supplied path.
        if ([IO.Directory]::Exists($folder)) {
            try { [IO.Directory]::Delete($folder, $true) } catch { Write-Warning 'پاک‌سازی فایل‌های موقت گزارش کامل نشد.' }
        }
    }
}
