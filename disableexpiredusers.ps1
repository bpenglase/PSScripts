# Find and disable computers that either haven't been logged in, or fall past the $Cutoff date
# Script heavily modified from http://powershell.nicoh.me/powershell-1/active-directory/disable-and-move-inactive-computer-accounts
# Other Resources:
#	http://www.tomsitpro.com/articles/active_directory-powershell-windows_server-aduc-scripting,2-250.html
#	https://mikegriffin.ie/blog/20120521-filtering-with-powershell-and-get-adcomputer/

$RootDomain="DC=example,DC=org"
$CutOff = (Get-Date).AddDays(0)
$UsersDisables = "OU=School,OU=Disabled,$RootDomain"
$UsersBase = "OU=School,OU=Students,$RootDomain"
$Users = Get-ADUser -SearchBase $UsersBase -filter { AccountExpirationDate -lt $Today }
ForEach ($User1 in $Users) {
	if ($User1 -ne $null) {
		$Name = $User1.SamAccountName
		$User = Get-ADUSer $name -prop Description | Select DistinguishedName,Description
		#Disable User Object
		Write-Host "Disabling: $name" -Fore "Red"
		Set-ADUser $name -Enabled $false
		#Update Description
		Write-Host "Updating Description with Source OU and Date"
		$New_Desc = ($User).Description + " - Moved From: " + ($User).DistinguishedName + " On " + (Get-Date -format d)
		Set-ADUser -identity $User.DistinguishedName -Description $New_Desc
		#Move User to Disabled OU
		Write-Host "Moving User account $name to $UsersDisables"
		Get-ADUser $name | Move-ADObject -TargetPath $UsersDisables
	}
}
