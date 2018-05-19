$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

[String] $InsuredName = ""
[String] $ClaimNumber = ""
[DateTime] $DateOfLoss = Get-Date
[DateTime] $AssignmentDate = Get-Date
[String] $ClaimName = ""

. (Join-Path $PSScriptRoot "INI.ps1")

$ClaimInitConfig = Get-IniContent (Join-Path $PSScriptRoot "ClaimInit.ini")

Function Initialize-Claim([String] $InsuredName, [String] $ClaimNumber, [DateTime] $DateOfLoss, [DateTime] $AssignmentDate) {
    $Script:InsuredName = $InsuredName
    $Script:ClaimNumber = $ClaimNumber
    $Script:DateOfLoss = $DateOfLoss
    $Script:AssignmentDate = $AssignmentDate
	Write-Host ("{0}|{1}|{2}|{3}" -f $InsuredName, $ClaimNumber, $DateOfLoss, $AssignmentDate)
	Write-Host ("{0}|{1}|{2}|{3}" -f $Script:InsuredName, $Script:ClaimNumber, $Script:DateOfLoss, $Script:AssignmentDate)
}

Function Split-Name([String] $Name) {
   Return $Name -split ", "
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
	Try {
		$ClaimDriveCredentials = New-Object System.Management.Automation.PSCredential ($Script:ClaimInitConfig["CREDENTIALS"]["USERNAME"], (ConvertTo-SecureString $Script:ClaimInitConfig["CREDENTIALS"]["PASSWORD"]))
		New-PSDrive "T" -PSProvider FileSystem -root ($Script:ClaimInitConfig["CONFIG"]["BASECLAIMFOLDER"]) -credential $ClaimDriveCredentials
	    $ClaimFolderName = Join-Path "T:\" (Get-ClaimName)
	    If (Test-Path $ClaimFolderName) {
	        Throw ("The claim folder '{0}' already exists.  Enter new data and try again." -f $ClaimFolderName)
	        Return $False 
	    } Else {
	        $null = New-Item $ClaimFolderName -Type directory
	        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd} REFERENCE" -f $Script:AssignmentDate)) -Type directory
	        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd}a FNOL" -f $Script:AssignmentDate)) -Type directory
	        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd}b ACK" -f $Script:AssignmentDate)) -Type directory
	        $null = New-Item (Join-Path $ClaimFolderName ("{0:yyyy.MM.dd}c INSP IMAGES" -f $Script:AssignmentDate)) -Type directory
	        $null = New-Item (Join-Path $ClaimFolderName ("x{0:yyyy.MM.dd} 1ST RPT" -f ($Script:AssignmentDate.AddDays(7)))) -Type directory
	        $null = New-Item (Join-Path $ClaimFolderName ("x{0:yyyy.MM.dd} 2ND RPT" -f ($Script:AssignmentDate.AddDays(30)))) -Type directory
           $null = New-Item (Join-Path $ClaimFolderName ("x{0:yyyy.MM.dd} DIARY" -f ($Script:AssignmentDate.AddDays(40)))) -Type directory
           $null = Copy-Item (Join-Path $PSScriptRoot "Template Notepad.docx") (Join-Path $ClaimFolderName ("01. {0} Notepad.docx" -f (Split-Name $Script:InsuredName)))
	        Return $True
	   }
	   Remove-PSDrive "T"
   } Catch {
   		Throw $_.Exception
   }
}

Function Create-ClaimReminder() {
	$olAppointmentItem = 1
    $outlookApplication = New-Object -ComObject 'Outlook.Application'
    #$outlookCalendar = $outlookApplication.Session.Folders.Item(4).Folders.Item(1)
    $newCalenderItem = $outlookApplication.CreateItem($olAppointmentItem)

    $newCalenderItem.Start = $Script:AssignmentDate.AddDays(7)
    $newCalenderItem.AllDayEvent = $True 
    $newCalenderItem.Subject = Get-AppointmentName
    $newCalenderItem.Body = ("Reminder to submit reports for the claim '{0}'" -f (Get-ClaimName))
    $newCalenderItem.ReminderMinutesBeforeStart = 15
    $newCalenderItem.ReminderSet = $True

    $newCalenderItem.Save()
}



