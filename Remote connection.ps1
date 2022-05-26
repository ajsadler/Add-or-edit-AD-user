$logonName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$logonNameSplit = $logonName.Split("\")
$logonUsername = $logonNameSplit[1]

$admUsername = "HGROUP\" + $logonUsername + "-adm"
$cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $admUsername
$cred
New-PSSession 