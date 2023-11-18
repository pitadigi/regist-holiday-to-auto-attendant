Import-Module MicrosoftTeams

Connect-MicrosoftTeams

$japanHoliday = @{}
$holidays = Get-CsOnlineSchedule | Where-Object {$_.Type -eq "Fixed"}
foreach ($holiday in $holidays) {
    if ($holiday.Name -eq "日本の祝日") {
        $japanHoliday = $holiday
    }
}

$rows = Import-Csv -Encoding UTF8 -Path ".\syukujitsu.csv"

$thisYear = Get-Date -UFormat %Y
foreach ($row in $rows) {
    $holidayYear = ($row."国民の祝日・休日月日" | Get-Date -UFormat %Y)

    if ($holidayYear -ge $thisYear) {
        $date = $row."国民の祝日・休日月日" | Get-Date -UFormat %d/%m/%Y
        If ($teamsSchedule[$holiday] -ne $date) {
            $holiday = $row."国民の祝日・休日名称"

            Write-Host ("Processing {0} on {1}" -f $holiday, $date)

            $holidayDateRange = New-CsOnlineDateTimeRange -Start $date
            $japanHoliday.FixedSchedule.DateTimeRanges += $holidayDateRange
        }
    }
}

Set-CsOnlineSchedule -Instance $japanHoliday
