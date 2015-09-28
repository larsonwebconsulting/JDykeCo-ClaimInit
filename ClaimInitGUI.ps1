$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Not running with administrative rights. Attempting to elevate..."
    If (Test-Path (Join-Path $PSScriptRoot "debug.txt")) {
    	$command = "-noexit"
    } Else {
    	$command = "-WindowStyle Hidden"
    }
    $command = "$command -STA -NoProfile -ExecutionPolicy bypass -command &'$PSScriptRoot\ClaimInitGUI.ps1'"
    Start-Process powershell -verb runas -argumentlist $command
    Exit
}

Function Show-MessageBox([String] $Message, [String] $Title = "Message", [Int] $BoxType = 0, [int] $Icon = 64) {
	# 0:	OK
	# 1:	OK Cancel
	# 2:	Abort Retry Ignore
	# 3:	Yes No Cancel
	# 4:	Yes No
	# 5:	Retry Cancel
	
	# None 0
	# Hand 16
	# Error 16
	# Stop 16
	# Question 32
	# Exclamation 48
	# Warning 48
	# Asterisk 64
	# Information 64

	[System.Windows.Forms.MessageBox]::Show($Message, $Title, $BoxType, $Icon)
}

Function Show-Error([String] $Message) {
	Show-MessageBox -Message $Message -Title "Error" -BoxType 0 -Icon 16
}

Function Show-Warning([String] $Message) {
	Show-MessageBox -Message $Message -Title "Warning" -BoxType 0 -Icon 48
}

Function Show-Information([String] $Message) {
	Show-MessageBox -Message $Message -Title "Information" -BoxType 0 -Icon 64
}

#Get Script path and name
$Script = $MyInvocation.MyCommand.Definition

If ($PSVersionTable.CLRVersion.Major -ne 4) {
	Show-Error ".NET Framework 4.0 is required for this utility.  Please execute the Enable-DotNet4Access.cmd as an administrator and try again." | Out-Null
	Exit
}

#Validate that Script is launched
$IsSTAEnabled = $host.Runspace.ApartmentState -eq 'STA'
If ($IsSTAEnabled -eq $false) {
    "Script is not running in STA mode. Switching to STA Mode..."

    #Launch script in a separate PowerShell process with STA enabled
    Start-Process powershell.exe -ArgumentList "-STA -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command $Script"
    Exit
}

. (Join-Path $PSScriptRoot "ClaimInit.ps1")

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window x:Class="jdykecoClaimGui.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        mc:Ignorable="d"
        Title="J. Dyke &amp; Co. Claim Initiation" Height="223.77" Width="306.664" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Grid>
        <Image x:Name="JDykeLogo_round_jpg" Source="C:\\jdcbot\\JDykeLogo.round.jpg" Stretch="Fill" Opacity="0.10"/>
        <Label x:Name="lblInsuredName" Content="Insured Name:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblClaimNumber" Content="Claim Number:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblDateOfLoss" Content="Date of Loss:" HorizontalAlignment="Left" Margin="10,67,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblAssignmentDate" Content="Assignment Date:" HorizontalAlignment="Left" Margin="10,93,0,0" VerticalAlignment="Top"/>
        <DatePicker x:Name="dttmDateOfLoss" HorizontalAlignment="Left" Margin="114,67,0,0" VerticalAlignment="Top" RenderTransformOrigin="1.455,0.638" Width="166"/>
        <DatePicker x:Name="dttmAssignmentDate" HorizontalAlignment="Left" Margin="114,96,0,0" VerticalAlignment="Top" RenderTransformOrigin="1.455,0.638" Width="166"/>
        <TextBox x:Name="txtInsuredName" HorizontalAlignment="Left" Height="23" Margin="114,12,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="166"/>
        <TextBox x:Name="txtClaimNumber" HorizontalAlignment="Left" Height="23" Margin="114,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="166"/>
        <Button x:Name="btnOK" Content="OK" HorizontalAlignment="Left" Margin="10,160,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="btnCancel" Content="Exit" HorizontalAlignment="Left" Margin="210,160,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="btnClear" Content="Clear" HorizontalAlignment="Left" Margin="114,160,0,0" VerticalAlignment="Top" Width="75"/>
        <CheckBox x:Name="chkCreateReminder" Content="Create Reminder Appointment" HorizontalAlignment="Left" Margin="15,127,0,0" VerticalAlignment="Top" IsEnabled="True" IsChecked="True"/>
    </Grid>
</Window>
"@   
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
try{
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML

    #Read XAML     
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 

    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}catch{
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    $_.Exception
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
	If ($global:ReadmeDisplay -ne $true) {
		Write-Host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow
		$global:ReadmeDisplay=$true
	}
	Write-Host "Found the following interactable elements from our form" -ForegroundColor Cyan
	Get-Variable WPF*
}
 
Get-FormVariables
 
#===========================================================================
# Actually make the objects work
#===========================================================================

Function Clear-Values() {
	$WPFdttmAssignmentDate.SelectedDate = Get-Date
	$WPFdttmDateOfLoss.Text = ""
	$WPFtxtInsuredName.Text = ""
	$WPFtxtClaimNumber.Text = ""
}

$WPFbtnClear.Add_Click({
	Clear-Values
})
 
$WPFbtnOK.Add_Click({
	Write-Host ("{0}|{1}|{2}|{3}" -f $WPFtxtInsuredName.Text, $WPFtxtClaimNumber.Text, $WPFdttmDateOfLoss.SelectedDate, $WPFdttmAssignmentDate.SelectedDate)
    Initialize-Claim $WPFtxtInsuredName.Text $WPFtxtClaimNumber.Text $WPFdttmDateOfLoss.SelectedDate $WPFdttmAssignmentDate.SelectedDate
    
    Try {
	    If ((Create-ClaimFolder) -And ($WPFchkCreateReminder.IsChecked)){
	        Create-ClaimReminder
	    }
	} Catch {
		Show-Error -Message $_.Exception.Message 
	}
})

$WPFbtnCancel.Add_Click({
    $Form.Close()
}) 
#===========================================================================
# Shows the form
#===========================================================================
#write-host "To show the form, run the following" -ForegroundColor Cyan
Clear-Values
$Form.ShowDialog() | Out-Null
