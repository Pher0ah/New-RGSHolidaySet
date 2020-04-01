<#
.SYNOPSIS
    Creates the required Holidays in Lync/SfB

.DESCRIPTION
    This script will use a Holiday file in the current folder to create the Holidays according
    to the Holiday List given in the command-line .

.FORMAT
    The file format should be in CSV with two columns only:
        Name             A Name for the specific holiday
        Date             The date of the holiday in DD/MM/YYYY format

.NOTES
    Version          : 1.0
    Required Infra.  : Lync Server 2013/SfB with Enterprise Voice

    Last Updated     : 04/09/2014
    Wishlist         : Let me know and I will add it
    Known Issues     : none so far, your feedback is always welcome

    Acknowledgements : -

    Author(s)        : Hany Elkady (pheroah@gmail.com)
    Website          : http://pher0ah.blogspot.com.au

    Disclaimer       : The rights to the use of this script is give as is without warranty
                       of any kind implicit or otherwise

.LINK
    http://pher0ah.blogspot.com.au/XXXXX

.PARAMETER HolidaySet
    Name of the Holiday List to either create or update

.EXAMPLE
    .\New-RGSHolidays.ps1 VIC-Holidays

.INPUTS
	HolidaySet Names can be piped from another command if required.
#>
#Requires -Version 2.0

#---------------------------------------------------------------------
#region Script Parameters
#
#---------------------------------------------------------------------
[CmdletBinding(SupportsShouldProcess=$True)]
param(
    # Define HolidayList to use
    [parameter(Mandatory = $True)]
    [alias("d")][string] $HolidaySet
        
)
#endregion Script Parameters
#---------------------------------------------------------------------

#In case someone runs this from a non-Lync shell
Import-Module Lync

$HolidaysFile = "$HolidaySet.csv"
$PoolFQDN = (get-cspool).Identity[0]
$Holidays = @()
$NewHolidays = Import-Csv .\$HolidaysFile
$Parent = "ApplicationServer:$PoolFQDN"

foreach($Holiday in $NewHolidays){
    $HolidayName = $Holiday.Name
    $HolidayStartDate = ($Holiday.Date)+" 12:00:00 AM"
    $HolidayEndDate = ($Holiday.Date)+" 11:59:59 PM"
    write-Host $HolidayName  $HolidayStartDate  $HolidayEndDate
    $Holidays += (New-CsRgsHoliday -Name $HolidayName -StartDate $HolidayStartDate -EndDate $HolidayEndDate)
}


$theSet = (Get-CSRGSHolidaySet -Identity $Parent |?{$_.Name -eq $HolidaySet})

If($theSet){
    $theSet.HolidaySet.Add($Holidays)
    Set-CsRgsHolidaySet -Instance $theSet
}else{  
    New-CsRgsHolidaySet -Parent $Parent -Name $HolidaySet -HolidayList($Holidays)
}