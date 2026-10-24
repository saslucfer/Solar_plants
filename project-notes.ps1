function Read-ProjectNotes([string]$folder) {
    $path = Join-Path $folder 'project-notes.json'
    if (-not [IO.File]::Exists($path)) { return [pscustomobject]@{ schemaVersion = 1; notes = @() } }
    try {
        $store = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($store.schemaVersion -ne 1 -or $null -eq $store.notes -or $store.notes -isnot [array]) { throw 'Invalid notes format' }
        $ids = @{}
        foreach ($note in $store.notes) {
            if (-not $note.id -or $ids.ContainsKey([string]$note.id) -or $note.body -isnot [string] -or [int]$note.revision -lt 1) { throw 'Invalid note record' }
            $ids[[string]$note.id] = $true
        }
        return $store
    }
    catch { throw '[500] فایل یادداشت‌های پروژه قابل خواندن نیست. نسخه پشتیبان را بررسی کنید؛ فایل اصلی بازنویسی نشد.' }
}

function Save-ProjectNote([string]$folder, $request) {
    try { $lock = [IO.File]::Open((Join-Path $folder 'project-notes.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
    catch { throw '[423] یادداشت‌ها در حال ذخیره در اجرای دیگری از داشبورد هستند؛ دوباره تلاش کنید.' }
    try { return Save-ProjectNoteCore $folder $request }
    finally { $lock.Dispose() }
}

function Save-ProjectNoteCore([string]$folder, $request) {
    $projectIds = @('ali-abad', 'damghan', 'evaz', 'hosnije', 'lar', 'marave-tape', 'natanz', 'safi-abad', 'shahreza')
    if (-not $request -or $projectIds -notcontains [string]$request.plantId) { throw '[400] نیروگاه معتبر نیست.' }
    if (@('note', 'comment', 'report') -notcontains [string]$request.kind) { throw '[400] نوع یادداشت معتبر نیست.' }
    if ($request.body -isnot [string] -or [string]::IsNullOrWhiteSpace($request.body) -or $request.body.Length -gt 6000) { throw '[400] متن توضیحات باید بین ۱ تا ۶۰۰۰ نویسه باشد.' }
    if ($null -ne $request.title -and ($request.title -isnot [string] -or $request.title.Length -gt 140)) { throw '[400] عنوان باید حداکثر ۱۴۰ نویسه باشد.' }
    $store = Read-ProjectNotes $folder
    $now = [DateTime]::UtcNow.ToString('o')
    $existing = $null
    if ($request.id) {
        $existing = $store.notes | Where-Object id -eq ([string]$request.id) | Select-Object -First 1
        if (-not $existing) { throw '[404] یادداشت پیدا نشد. فهرست را تازه کنید.' }
        if ($existing.plantId -ne [string]$request.plantId) { throw '[400] نیروگاه یادداشت قابل تغییر نیست.' }
        if ($request.expectedRevision -isnot [int] -and $request.expectedRevision -isnot [long]) { throw '[400] نسخه یادداشت معتبر نیست.' }
        if ([long]$request.expectedRevision -ne [long]$existing.revision) { throw '[409] این یادداشت در صفحه دیگری ویرایش شده است. متن شما حفظ شد؛ فهرست را تازه کنید و دوباره ویرایش را آغاز کنید.' }
    }
    $record = [pscustomobject][ordered]@{
        id = if ($existing) { $existing.id } else { [Guid]::NewGuid().ToString() }
        plantId = [string]$request.plantId
        kind = [string]$request.kind
        title = ([string]$request.title).Trim()
        body = $request.body.Trim()
        createdAt = if ($existing) { $existing.createdAt } else { $now }
        updatedAt = $now
        revision = if ($existing) { [int]$existing.revision + 1 } else { 1 }
    }
    if ($existing) {
        $store.notes = @($store.notes | ForEach-Object { if ($_.id -eq $existing.id) { $record } else { $_ } })
    } else { $store.notes = @($store.notes) + @($record) }
    $path = Join-Path $folder 'project-notes.json'
    $temporaryPath = Join-Path $folder ('project-notes.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    try {
        $json = ConvertTo-Json -InputObject $store -Depth 8
        [IO.File]::WriteAllText($temporaryPath, $json, [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($path)) { [IO.File]::Replace($temporaryPath, $path, ($path + '.bak')) }
        else { [IO.File]::Move($temporaryPath, $path) }
    }
    finally { if ([IO.File]::Exists($temporaryPath)) { [IO.File]::Delete($temporaryPath) } }
    return [ordered]@{ note = $record }
}
