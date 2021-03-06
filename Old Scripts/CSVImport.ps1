# Bring in the AD CMDLETS
import-module activedirectory
###############################
# Select-FileDialog Function  #
# Created by Hugo Peeters     #
# http://www.peetersonline.nl #
###############################

# Note: store in your profile for easy use
# Example use:
# $file = Select-FileDialog -Title "Select a file" -Directory "D:\scripts" -Filter "Powershell Scripts|(*.ps1)"
function Select-FileDialog
{
	param([string]$Title,[string]$Directory,[string]$Filter="Comma-Separated Values (*.csv)|*.csv|All Files (*.*)|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	Else
	{
		Write-Error "Operation cancelled by user."
	}
}

# Call Function to display file selection dialog, and set the variable to the resulting path. 
$fileselect=Select-FileDialog
write-host $fileselect

# Import the CSV File
try {
    $file = import-csv $fileselect
       write-host "Imported"
}
catch {
    write-host File Import Failed.
    exit
}

# Create Empty Arrays
$hsstudents=@()
$msstudents=@()
$uesstudents=@()
$lesstudents=@()

# Begin to parse through
foreach ($user in $file) {
    # Grab and format some variables
    $firstname = $user."First Name ".trim()
    $firststrip = $firstname -replace " ",""
    $firststrip = $firststrip -replace ".``",""
    $firststrip = $firststrip -replace ".`'",""
    $FirstInitial = $firstname.substring(0,1)
    $lastname = $user."Last Name ".Trim()
    $laststrip = $lastname -replace " ",""
    $laststrip = $laststrip -replace ".``",""
    $laststrip = $laststrip -replace ".`'",""
    $oldstyle = $laststrip+$Firstinitial
    $displayname = $lastname+", "+$firstname
    $samaccountname = $laststrip+$firststrip
    $logonname = $laststrip+$firststrip
    $id = $user.'Student ID '.trim()
    # If the SAMAccountName ends up over 20, we have to shorten it for 'backwards compatibility'
    if ($samaccountname.Length -gt 20) {
        $samaccountname = $samaccountname.Substring(0,20)
    }
    $domain = "@example.org"
    $studentdomain = "@students.example.org"
    $fileserver = "\\fs1.example.org\students\"
    $gradyear = $user.'Graduation Year '.trim()
    # Figure out what building the students are in based on graduation year
    switch($gradyear) {
        2014 {$building="hs"} #Senior
        2015 {$building="hs"} #Junior
        2016 {$building="hs"} #Sophmore
        2017 {$building="hs"} #Freshman
        2018 {$building="ms"} #Eigth
        2019 {$building="ms"} #Seventh
        2020 {$building="ms"} #Sixth
        2021 {$building="ues"} #Fifth
        2022 {$building="ues"} #Fourth
        2023 {$building="ues"} #Third
        2024 {$building="les"} #Second
        2025 {$building="les"} #First
        2026 {$building="les"} #Kindergarten
        default {$building="Generic"} 
    }
    # Form the password based on used convention
    switch($building) {
        hs {$password="basepassword"+$id
            $changepass = 1
            $techperson = "hsteacher@example.org" }
        ms {$password="basepassword"+$id
            $changepass = 1
            $techperson = "msteacher@example.org" }
        ues {$password="basepassword"
            $changepass = 0
            $techperson = "uesteacher@example.org" }
        les {$password="basepassword"
            $changepass = 0
            $techperson = "lesteacher@example.org" }
        default {$password="basepassword"
                $changepass = 0
				$techperson = "it@example.org" } 
    }
    # Now form the OU for the students based on the building and gradyear.
    $ou="OU="+$gradyear+",OU="+$building+",OU=Students,dc=example,dc=org"
    try {
        # Check to see if the older format username exists
        # If it exists, say so, and continue on with the next user
        $user1 = get-aduser -identity $oldstyle
    }
    catch {
        # Old style does not exist, check for new style
        try {
            # If it exists, say so and continue on with the next user
            $user1 = Get-ADUser -identity $samaccountname
        }
        catch {
            # Neither old style or new style exists, Create the user
            try {
                write-host -------------------------------------------------
                write-host Creating in: $ou
                new-aduser -name $displayname `
                -samaccountname $samaccountname `
                -Surname $lastname `
                -GivenName $firstname `
                -UserPrincipalName $logonname$domain `
                -DisplayName $displayname `
                -path $ou `
                -Description $gradyear `
                -emailaddress $logonname$studentdomain `
                -HomeDirectory $fileserver$gradyear"\"$samaccountname `
                -HomeDrive "H:" `
                -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
                -ChangePasswordAtLogon $changepass `
                -Enabled $true
                write-host Creating Account for: $FirstName $LastName
                write-host Username: $samaccountname
                write-host Password: $password
                write-host email: $logonname$studentdomain
                # Sleep to let the account get added to the system
                # This has to be fairly high to let the sysytem propagate the new account,
                # As we don't know if we're connected to the same DC as the target fileserver.
                start-sleep -s 10
                Add-ADGroupMember -Identity $gradyear -Member $samaccountname 
                if (!(Test-Path -Path $fileserver$gradyear"\"$samaccountname )) {
                    write-host Creating HomeFolder
                    New-Item $fileserver$gradyear"\"$samaccountname -Type Directory -Force
                } Else {
                    write-host Home Folder Exists
                }
                $acl = Get-ACL $fileserver$gradyear"\"$samaccountname
            	$UserObj = New-Object System.Security.Principal.NTAccount($SamAccountName)
                $rights=[System.Security.AccessControl.FileSystemRights]::FullControl
                $inheritance=[System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
                $propagation=[System.Security.AccessControl.PropagationFlags]::InheritOnly
                $allowdeny=[System.Security.AccessControl.AccessControlType]::Allow
                $dirace=New-Object System.Security.AccessControl.FileSystemAccessRule ($samaccountname,$rights,$inheritance,$propagation,$allowdeny)
                $ACL.AddAccessRule($dirace)
                $rights=[System.Security.AccessControl.FileSystemRights]::FullControl
                $inheritance=[System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
                $propagation=[System.Security.AccessControl.PropagationFlags]::None
                $allowdeny=[System.Security.AccessControl.AccessControlType]::Allow
                $dirace=New-Object System.Security.AccessControl.FileSystemAccessRule ($samaccountname,$rights,$inheritance,$propagation,$allowdeny)
                $ACL.AddAccessRule($dirace)                
                $acl.SetOwner($UserObj)
            	Set-ACL $fileserver$gradyear"\"$samaccountname $acl
                switch($building) {
                    hs {$hsstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                    ms {$msstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                    ues {$uesstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                    les {$lesstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                }
                # Clean up Variables
                remove-variable acl
                remove-variable userobj
                remove-variable rights
                remove-variable inheritance
                remove-variable propagation
                remove-variable allowdeny
                remove-variable dirace
            } Finally {
                write-host -------------------------------------------------
            }
        }
    }
}

$smtpServer = "mail.example.org"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "accounts@example.org"
# Send to HS
if ($hsstudents.length -ge 1) {
    $msg.To.clear()
    $msg.body = ""
    $msg.subject = "New Accounts - HS"
    $msg.To.Add($hsstudents[0][0])
    $msg.To.Add("it@example.org")
    for ($i=0; $i -le $hsstudents.length-1; $i++) {
        $msg.body += "
Creating in: "+$hsstudents[$i][1]+"
Created Account for: "+$hsstudents[$i][2]+" "+$hsstudents[$i][3]+"
Username: "+$hsstudents[$i][4]+"
Password: "+$hsstudents[$i][5]+"
email: "+$hsstudents[$i][6]+$hsstudents[$i][7]+"

" 
    }
    $smtp.Send($msg)
}
# Send to MS
if ($msstudents.length -ge 1) {
    $msg.To.Clear()
    $msg.body = ""
    $msg.subject = "New Accounts - MS"
    $msg.To.Add($msstudents[0][0])
    $msg.To.Add("it@example.org")
    for ($i=0; $i -le $msstudents.length-1; $i++) {
        $msg.body += "
Creating in: "+$msstudents[$i][1]+"
Created Account for: "+$msstudents[$i][2]+" "+$msstudents[$i][3]+"
Username: "+$msstudents[$i][4]+"
Password: "+$msstudents[$i][5]+"
email: "+$msstudents[$i][6]+$msstudents[$i][7]+"

" 
    }
    $smtp.Send($msg)
}
# Send to UES
if ($uesstudents.length -ge 1) {
    $msg.To.Clear()
    $msg.body = ""
    $msg.subject = "New Accounts - UES"
    $msg.To.Add($uesstudents[0][0])
    $msg.To.Add("it@example.org")
    for ($i=0; $i -le $uesstudents.length-1; $i++) {
        $msg.body += "
Creating in: "+$uesstudents[$i][1]+"
Created Account for: "+$uesstudents[$i][2]+" "+$uesstudents[$i][3]+"
Username: "+$uesstudents[$i][4]+"
Password: "+$uesstudents[$i][5]+"
email: "+$uesstudents[$i][6]+$uesstudents[$i][7]+"

" 
    }
    $smtp.Send($msg)
}
# Send to LES
if ($lesstudents.length -ge 1) {
    $msg.To.Clear()
    $msg.body = ""
    $msg.subject = "New Account - LES"
    $msg.To.Add($lesstudents[0][0])
    $msg.To.Add("it@example.org")
    for ($i=0; $i -le $lesstudents.length-1; $i++) {
        $msg.body += "
Creating in: "+$lesstudents[$i][1]+"
Created Account for: "+$lesstudents[$i][2]+" "+$lesstudents[$i][3]+"
Username: "+$lesstudents[$i][4]+"
Password: "+$lesstudents[$i][5]+"
email: "+$lesstudents[$i][6]+$lesstudents[$i][7]+"

" 
    }
    $smtp.Send($msg)
}
