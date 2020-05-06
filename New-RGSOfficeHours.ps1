#In case someone runs this from a non-Lync shell
Import-Module Lync

#Create Working Hours
$HoursWeekDays = New-CSRgsTimeRange -Name "ESV Weekday Hours" -OpenTime "08:30" -CloseTime "16:30"
$HoursCOES = New-CSRgsTimeRange -Name "COES Weekday Hours" -OpenTime "09:00" -CloseTime "16:00"
$HoursLicensingMTF = New-CSRgsTimeRange -Name "Licensing Hours (Mon/Thu/Fri)" -OpenTime "09:00" -CloseTime "16:00"
$HoursLicensingTue = New-CSRgsTimeRange -Name "Licensing Hours (Tue)" -OpenTime "12:30" -CloseTime "16:00"
$HoursLicensingWed = New-CSRgsTimeRange -Name "Licensing Hours (Wed)" -OpenTime "09:00" -CloseTime "12:30"
$Hours247 = New-CSRgsTimeRange -Name "24x7 Hours" -OpenTime "00:00" -CloseTime "23:59"

$Pools = (Get-CsPool |?{$_.Services -match "ApplicationServer"}).Services -match 'ApplicationServer'

foreach ($AppPool in $Pools){
  New-CsRgsHoursOfBusiness -Parent $AppPool -Name "Gas Operations Hours" -MondayHours1 $HoursLicensingMTF `
                                                                    -TuesdayHours1 $HoursLicensingTue `
                                                                    -WednesdayHours1 $HoursLicensingWed `
                                                                    -ThursdayHours1 $HoursLicensingMTF `
                                                                    -FridayHours1 $HoursLicensingMTF `
                                                                    -SaturdayHours1 $Null `
                                                                    -SundayHours1 $Null
}
