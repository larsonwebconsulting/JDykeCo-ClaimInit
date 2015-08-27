[String] $InsuredName = ""
[String] $ClaimNumber = ""
[DateTime] $DateOfLoss = Get-Date
[DateTime] $AssignmentDate = Get-Date
[String] $ClaimName = ""

. .\INI.ps1

$ClaimInitConfig = Get-IniContent ".\ClaimInit.ini"
$ClaimInitConfig["CONFIG"]["BASECLAIMFOLDER"]

Function Initialize-Claim([String] $InsuredName, [String] $ClaimNumber, [DateTime] $DateOfLoss, [DateTime] $AssignmentDate) {
    $Script:InsuredName = $InsuredName
    $Script:ClaimNumber = $ClaimNumber
    $Script:DateOfLoss = $DateOfLoss
    $Script:AssignmentDate = $AssignmentDate
	Write-Host ("{0}|{1}|{2}|{3}" -f $InsuredName, $ClaimNumber, $DateOfLoss, $AssignmentDate)
	Write-Host ("{0}|{1}|{2}|{3}" -f $Script:InsuredName, $Script:ClaimNumber, $Script:DateOfLoss, $Script:AssignmentDate)
}

Function Get-ClaimName() {
    $Script:ClaimName = "{0}, {1}, {2:yyyyMMdd}" -f $Script:InsuredName, $Script:ClaimNumber, $Script:DateOfLoss
    Return $Script:ClaimName
}

Function Get-AppointmentName() {
	$NamePieces = $Script:InsuredName -split ","
	return $NamePieces[0] + ", " + $NamePieces[1]
}

Function Create-ClaimFolder() {
    $ClaimFolderName = Join-Path $Script:ClaimInitConfig["CONFIG"]["BASECLAIMFOLDER"] (Get-ClaimName)
    If (Test-Path $ClaimFolderName) {
        #Write-Error -Message ("The claim folder '{0}' already exists." -f $ClaimFolderName) -Category InvalidData -RecommendedAction "Enter new data and try again."
        Throw ("The claim folder '{0}' already exists.  Enter new data and try again." -f $ClaimFolderName)
        Return $False 
    } Else {
        $null = New-Item $ClaimFolderName -Type directory
        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd}a FNOL" -f $Script:AssignmentDate)) -Type directory
        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd}b ACK" -f $Script:AssignmentDate)) -Type directory
        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd}c INSP IMAGES" -f $Script:AssignmentDate)) -Type directory
        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd} 1ST REPORT" -f ($Script:AssignmentDate.AddDays(7)))) -Type directory
        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd} STATUS" -f ($Script:AssignmentDate.AddDays(38)))) -Type directory
        Return $True
   }
}

Function Create-ClaimReminder() {
    $outlookApplication = New-Object -ComObject 'Outlook.Application'
    #$outlookCalendar = $outlookApplication.Session.Folders.Item(4).Folders.Item(1)
    $newCalenderItem = $outlookApplication.CreateItem('olAppointmentItem')

    $newCalenderItem.Start = $Script:AssignmentDate.AddDays(7)
    $newCalenderItem.AllDayEvent = $True 
    $newCalenderItem.Subject = Get-AppointmentName
    $newCalenderItem.Body = ("Reminder to submit reports for the claim '{0}'" -f (Get-ClaimName))
    $newCalenderItem.ReminderMinutesBeforeStart = 15
    $newCalenderItem.ReminderSet = $True

    $newCalenderItem.Save()
}



