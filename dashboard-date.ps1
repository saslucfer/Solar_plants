function Get-PersianDashboardDate {
    $tehran = [TimeZoneInfo]::FindSystemTimeZoneById('Iran Standard Time')
    $today = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tehran)
    $calendar = [Globalization.PersianCalendar]::new()
    $value = '{0:D4}/{1:D2}/{2:D2}' -f $calendar.GetYear($today), $calendar.GetMonth($today), $calendar.GetDayOfMonth($today)
    for ($digit = 0; $digit -lt 10; $digit++) {
        $value = $value.Replace([string]$digit, [string][char](0x06F0 + $digit))
    }
    return $value
}
