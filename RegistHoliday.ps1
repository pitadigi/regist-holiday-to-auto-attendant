Import-Module MicrosoftTeams

$rows = Import-Csv -Encoding Default -Path "\\Mac\Home\Downloads\syukujitsu.csv"

$holidayList = @()

$thisYear = Get-Date -UFormat %Y
foreach ($row in $rows) {
    $holidayYear = ($row."国民の祝日・休日月日" | Get-Date -UFormat %Y)

    if ($holidayYear -ge $thisYear) {
        $holiday = [PSCustomObject]@{}
        $startDateTime = @()
        for ($i=0 ; $i -lt $holidayList.Count ; $i++) {
            if ($holidayList[$i].Title -eq $row."国民の祝日・休日名称") {
                $startDateTime = @(
                    $holidayList[$i].StartDateTime1,
                    $holidayList[$i].StartDateTime2,
                    $holidayList[$i].StartDateTime3,
                    $holidayList[$i].StartDateTime4,
                    $holidayList[$i].StartDateTime5,
                    $holidayList[$i].StartDateTime6,
                    $holidayList[$i].StartDateTime7,
                    $holidayList[$i].StartDateTime8,
                    $holidayList[$i].StartDateTime9,
                    $holidayList[$i].StartDateTime10
                )
                for ($j=0 ; $j -lt $startDateTime.Count ; $j++) {
                    if ($startDateTime[$j] -eq "") {
                        $startDateTime[$j] = $row."国民の祝日・休日月日" | Get-Date -UFormat %d/%m/%Y
                        break
                    }
                }
                $holidayList[$i].StartDateTime1 = $startDateTime[0]
                $holidayList[$i].StartDateTime2 = $startDateTime[1]
                $holidayList[$i].StartDateTime3 = $startDateTime[2]
                $holidayList[$i].StartDateTime4 = $startDateTime[3]
                $holidayList[$i].StartDateTime5 = $startDateTime[4]
                $holidayList[$i].StartDateTime6 = $startDateTime[5]
                $holidayList[$i].StartDateTime7 = $startDateTime[6]
                $holidayList[$i].StartDateTime8 = $startDateTime[7]
                $holidayList[$i].StartDateTime9 = $startDateTime[8]
                $holidayList[$i].StartDateTime10 = $startDateTime[9]

                break
            }
        }
        if ($startDateTime.Count -eq 0) {
            $holiday = [PSCustomObject]@{
                Title = $row."国民の祝日・休日名称";
                StartDateTime1 = $row."国民の祝日・休日月日" | Get-Date -UFormat %d/%m/%Y;
                EndDateTime1 = "";
                StartDateTime2 = "";
                EndDateTime2 = "";
                StartDateTime3 = "";
                EndDateTime3 = "";
                StartDateTime4 = "";
                EndDateTime4 = "";
                StartDateTime5 = "";
                EndDateTime5 = "";
                StartDateTime6 = "";
                EndDateTime6 = "";
                StartDateTime7 = "";
                EndDateTime7 = "";
                StartDateTime8 = "";
                EndDateTime8 = "";
                StartDateTime9 = "";
                EndDateTime9 = "";
                StartDateTime10 = "";
                EndDateTime10 = ""
            }
            $holidayList += $holiday
        }

    }
}

$holidayListFile = "\\Mac\Home\Downloads\HolidayList.csv"
$holidayList | Export-Csv -Path $holidayListFile -Encoding UTF8

$Credential = Get-Credential
$teams = Connect-MicrosoftTeams -Credential $Credential

$bytes = [System.IO.File]::ReadAllBytes($holidayListFile)
$aa = Get-CsAutoAttendant
Import-CsAutoAttendantHolidays -Input $bytes -Identity $aa.Identity
