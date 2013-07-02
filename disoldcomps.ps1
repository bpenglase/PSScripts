# Find and disable computers that either haven't been logged in, or fall past the $Cutoff date
# Script heavily modified from http://powershell.nicoh.me/powershell-1/active-directory/disable-and-move-inactive-computer-accounts
# Other Resources:
#	http://www.tomsitpro.com/articles/active_directory-powershell-windows_server-aduc-scripting,2-250.html
#	https://mikegriffin.ie/blog/20120521-filtering-with-powershell-and-get-adcomputer/


$RootDomain="DC=example,DC=org"
$CutOff = (Get-Date).AddDays(-365)
$CompsDisabled = "OU=Computers,OU=Disabled,$RootDomain"
$CompsBase = "CN=Computers,$RootDomain"
$Machines = Get-ADComputer -SearchBase $CompsBase -filter * -prop LastLogonDate | where {($_.LastLogonDate -lt $cutoff) -or ($_.LastLogonDate -eq $null)}
ForEach ($Machine in $Machines) {
	$Name = $Machine.name
	# Ping machine to see if it's up, if not, continue
	$PingStatus = gwmi Win32_PingStatus -Filter "Address = '$name'" | Select-Object StatusCode
	if ($PingStatus.StatusCode -ne 0) {
		Write-Host $Name " : Offline" -Fore "Red"
		$Comp = Get-ADComputer $name -prop Description | Select DistinguishedName,Description
		#Disable Computer Object
		Write-Host "Disabling: $name" -Fore "Red"
		Set-ADComputer $name -Enabled $false
		#Update Description
		Write-Host "Updating Description with Source OU and Date"
		$New_Desc = ($Comp).Description + " - Moved From: " + ($Comp).DistinguishedName + " On " + (Get-Date -format d)
		Set-ADComputer -identity $Comp.DistinguishedName -Description $New_Desc
		#Move computer to Disabled OU
		Write-Host "Moving computer account $name to $CompsDisabled"
		Get-ADComputer $name | Move-ADObject -TargetPath $CompsDisabled
	}
}
