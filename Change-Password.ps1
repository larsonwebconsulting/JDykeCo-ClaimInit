. .\INI.ps1
$ClaimInitConfig = Get-IniContent ".\ClaimInit.ini"
$credential = Get-Credential
$ClaimInitConfig["CREDENTIALS"]["USERNAME"] = $credential.getNetworkCredential().UserName
$ClaimInitConfig["CREDENTIALS"]["PASSWORD"] = ($credential.getNetworkCredential().SecurePassword | ConvertFrom-SecureString)
Out-IniFile -InputObject $ClaimInitConfig -FilePath ".\ClaimInit.ini" -Force
