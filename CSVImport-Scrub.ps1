# Require PowerShell Version 3
#Requires -Version 3
# Import MSOnline PSCmdLets
##Requires -Modules msonline
# Bring in the AD CMDLETS
#Requires -Modules activedirectory

# Get O365 Credentials before we even get the file.. if we can't auth, we shouldn't create.
# Prompt for credentials and connect to O365 if not already connected
#Try {
#    Get-MsolDomain -ErrorAction Stop > $null
#} Catch {
#    try {
#        $Credentials = Get-Credential -Message "Enter Office 365 Credentials" -UserName "admin@domain.onmicrosoft.com"
#        connect-msolservice -Credential $Credentials
#    } Catch {
#        Write-Host -ForegroundColor Red Sorry, we need credentials to continue
#        Exit
#    }
#}
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
$fileselect=Select-FileDialog "Select CSV" "\\domain.org\Shares\Technology"
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

# Setup the disabled services
#$LicenseOptions = New-MsolLicenseOptions -AccountSkuId 'domain:STANDARDWOFFPACK_STUDENT' -DisabledPlans Yammer_EDU,Exchange_S_Standard,MCOStandard

# Create Empty Arrays
$hsstudents=@()
$msstudents=@()
$uesstudents=@()
$lesstudents=@()

# Begin to parse through
foreach ($user in $file) {
    # Grab and format some variables
    $firstname = $user."Student First Name".trim()
    $firststrip = $firstname -replace " ",""
    $firststrip = $firststrip -replace "``",""
    $firststrip = $firststrip -replace "`'",""
    $FirstInitial = $firstname.substring(0,1)
    $lastname = $user."Student Last Name".Trim()
    $laststrip = $lastname -replace " ",""
    $laststrip = $laststrip -replace "``",""
    $laststrip = $laststrip -replace "`'",""
    $middlename = $user."Student Middle Name".Trim()
    $middlestrip = $middlename -replace " ",""
    $middlestrip = $middlestrip -replace "``",""
    $middlestrip = $middlestrip -replace "`'",""
    $oldstyle = $laststrip+$Firstinitial
    $displayname = $lastname+", "+$firstname
    $displaynamealt = $lastname+", "+$firstname+" "+$middlename
    $samaccountname = $laststrip+$firststrip
    $samaccountnamealt = $laststrip+$firststrip+$middlestrip
    $logonname = $laststrip+$firststrip
    $logonnamealt = $laststrip+$firststrip+$middlestrip
    $id = $user.'Student Id'.trim()
    # If the SAMAccountName ends up over 20, we have to shorten it for 'backwards compatibility'
    if ($samaccountname.Length -gt 20) {
        $samaccountname = $samaccountname.Substring(0,20)
    }
    if ($samaccountnamealt.Length -gt 20) {
        $samaccountnamealt = $samaccountnamealt.Substring(0,20)
    }
    $domain = "@domain.org"
    $studentdomain = "@studentdomain.org"
    $fileserver = "\\domain.org\Users\students\"
    $gradyear = $user.'Grad Year'.trim()
    # Figure out what building the students are in based on graduation year
    switch($gradyear) {
        2018 {$building="hs"} #Senior
        2019 {$building="hs"} #Junior
        2020 {$building="hs"} #Sophmore
        2021 {$building="hs"} #Freshman
        2022 {$building="ms"} #Eigth
        2023 {$building="ms"} #Seventh
        2024 {$building="ms"} #Sixth
        2025 {$building="ues"} #Fifth
        2026 {$building="ues"} #Fourth
        2027 {$building="ues"} #Third
        2028 {$building="les"} #Second
        2029 {$building="les"} #First
        2030 {$building="les"} #Kindergarten
        default {$building="Generic"} 
    }
    # Form the password based on used convention
    switch($building) {
        hs {$password="meow"+$id
            $changepass = 1
            $techperson = "hslibrarian@domain.org, libassistant@domain.org"
            $homedirectory = $Null
            $homedrive = $null }
        ms {$password="meow"+$id
            $changepass = 1
            $techperson = "mslibrarian@domain.org" 
            $homedirectory = $Null
            $homedrive =$null }
        ues {$password="meow"+$id
            $changepass = 0
            $techperson = "ueslibrarian@domain.org" 
            $homedirectory = $fileserver+$gradyear+"\"+$samaccountname
            $homedrive =  "H:" }
        les {$password="meowmeow"
            $changepass = 0
            $techperson = "leslibrarian@domain.org" 
            $homedirectory = $fileserver+$gradyear+"\"+$samaccountname
            $homedrive =  "H:"}
        default {$password="meowmeow"
                $changepass = 0}
    }
    # Now form the OU for the students based on the building and gradyear.
    $ou="OU="+$gradyear+",OU="+$building+",OU=Students,OU=Users,OU=domain Objects,dc=domain,dc=org"
    # Default to no creation!
    $oktocreate = 2

    # Check to see if the user exists by ID number. This should be a match, the rest should be a just in case. 
    $userbyid = get-aduser -filter {EmployeeID -eq $id} -properties sn,givenname,enabled
    write-host $id
    if ($userbyid -ne $null) {
        # Need to code here to re-enable re-enrollees.
        write-host "User Found by ID"
        if ($userbyid.Enabled -eq $false) {
            Write-host "User found, but Disabled!!!!"
        }
        $oktocreate = 0
    }

    # If oktocreate is 0, we have a match above.. don't recreate. If anything else.. continue!
    if ($oktocreate -ne 0) {
        $useroldstyle = get-aduser -filter {samAccountName -eq $oldstyle} -Properties sn,givenname
        if ($useroldstyle -ne $null) {
            write-host "User Found as Oldstyle"
            if ($useroldstyle.sn -eq $lastname) {
                write-host "OldStyle: LastName Matches"
                if ($useroldstyle.GivenName -eq $firstname) {
                    write-host "Old Style: GivenName Matches - Looks like account exists, do not recreate"
                    $oktocreate = 0
                } else {
                    write-host Entire name does not match... OK to Create
                    $oktocreate = 1
                } # End Firstname
            } else {
                write-host "Oldstyle: Lastname didn't match for some reason, but we matched username. Allow creation"
                $oktocreate = 1
            } # end Lastname
        } else {
            write-host "Oldstyle: No matches. OK to Create."
            $oktocreate = 1
        } # End Oldstyle
    }
    
    # If oktocreate is 0, we have a match above.. don't recreate. If anything else.. continue!
    if ($oktocreate -ne 0) {
        $usernewstyle = get-aduser -filter {samAccountName -eq $samaccountname} -Properties sn,givenname,displayname
        if ($usernewstyle -ne $null) {
            write-host "User Found as NewStyle"
            if ($usernewstyle.sn -eq $lastname) {
                write-host "NewStyle: LastName Matches"
                if ($usernewstyle.GivenName -eq $firstname) {
                    write-host "NewStyle: GivenName Matches"
                    if ($usernewstyle.DisplayName -eq $displaynamealt) {
                        write-host "NewStyle: Display Name (Including Middle) Matched - Looks like account exists, Do not recreate - "+$usernewstyle
                        $oktocreate = 0
                    } else {
                        write-host "Entire name does not match... OK to Create."
                        # Since we got this far, it means normal conventions won't work (sharing same name, exlucding middle) Tell the script this by setting this to 3
                        $oktocreate = 3
                    } # End Display Name
                } else {
                    write-host "First name does not match... OK to create."
                    $oktocreate = 1
                } #End First Name
            } else {
                write-host "Newstyle: Lastname didn't match for some reason, but we matched username. Allow creation"
                $oktocreate = 1
            } # end Lastname
        } # end newstyle
    } # End ok -ne 0
    
    # Nothing blocking... create user
    if ($oktocreate -ne 0) {
        if ($oktocreate -eq 3) {
            # Swap in the alt name
            $displayname=$displaynamealt
            $samaccountname=$samaccountnamealt
            $logonname=$logonnamealt
        }
        try {
            write-host -------------------------------------------------
            write-host Creating in: $ou
            new-aduser -name $displayname `
                -samaccountname $samaccountname `
                -Surname $lastname `
                -GivenName $firstname `
                -UserPrincipalName $logonname$studentdomain `
                -DisplayName $displayname `
                -path $ou `
                -Description $gradyear `
                -emailaddress $logonname$studentdomain `
                -HomeDirectory $homedirectory `
                -HomeDrive $homedrive `
                -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
                -ChangePasswordAtLogon $changepass `
                -Enabled $true `
                -EmployeeID $id `
                -EmployeeNumber $id `
                -OtherAttributes @{EmployeeType="Student"}
            write-host Creating Account for: $displayname
            write-host Username: $samaccountname
            write-host Password: $password
            write-host email: $logonname$studentdomain
                
            # Sleep to let the account get added to the system
            # This has to be fairly high to let the sysytem propagate the new account,
            # As we don't know if we're connected to the same DC as the target fileserver, and to add group membership
            # Sleep a shorter time for those who don't have home directories
            if ($homedirectory -ne $null) {
                start-sleep -s 5
            } else {
                start-sleep -s 20
            }
            # Add User to their Gradyear Group
            Add-ADGroupMember -Identity $gradyear -Members $samaccountname 
            # Add User to the Google Printing Group (Enables them to use the Ricoh Cloud Printing)
            Add-ADGroupMember -Identity "Google Printing" -Members $samaccountname
            # Create HomeDirectory, if needed.
            if ($homedirectory -ne $null) {
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
                # Clean up Variables
                remove-variable acl
                remove-variable userobj
                remove-variable rights
                remove-variable inheritance
                remove-variable propagation
                remove-variable allowdeny
                remove-variable dirace 
            }
            switch($building) {
                hs {$hsstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                ms {$msstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                ues {$uesstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
                les {$lesstudents += (,($techperson,$ou,$firstname,$lastname,$samaccountname,$password,$logonname,$studentdomain))}
            }
            # Do Office 365 Licensing
            #write-host Setting Location
            #set-msoluser -UserPrincipalName $logonname$studentdomain -UsageLocation US
            #write-host Assigning O365 Education for Students
            #set-msoluserlicense -UserPrincipalName $logonname$studentdomain -AddLicenses 'domain:STANDARDWOFFPACK_STUDENT' -LicenseOptions $LicenseOptions
            #write-host Assigning O365 ProPlus for Students
            #set-msoluserlicense -UserPrincipalName $logonname$studentdomain -AddLicenses 'domain:OFFICESUBSCRIPTION_STUDENT'
        }
        Finally {
            write-host -------------------------------------------------
        }
    }
}

$smtpServer = "owa.domain.org"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "accounts@domain.org"
# Send to HS
if ($hsstudents.length -ge 1) {
    $msg.To.clear()
    $msg.body = ""
    $msg.subject = "New Accounts - HS"
    $msg.To.Add($hsstudents[0][0])
    $msg.To.Add("ITDept@domain.org")
    $msg.To.add("LMSAdmin@domain.org")
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
    $msg.To.Add("ITDept@domain.org")
    $msg.To.add("LMSAdmin@domain.org")
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
    $msg.To.Add("ITDept@domain.org")
    $msg.To.add("LMSAdmin@domain.org")
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
    $msg.To.Add("ITDept@domain.org")
    $msg.To.add("LMSAdmin@domain.org")
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

# Re-Generate Student Email Address Lists
get-aduser -searchbase "ou=2018,ou=hs,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\HS\Student Email Addresses\2018.csv" -NoTypeInformation
get-aduser -searchbase "ou=2019,ou=hs,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\HS\Student Email Addresses\2019.csv" -NoTypeInformation
get-aduser -searchbase "ou=2020,ou=hs,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\HS\Student Email Addresses\2020.csv" -NoTypeInformation
get-aduser -searchbase "ou=2021,ou=hs,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\HS\Student Email Addresses\2021.csv" -NoTypeInformation

get-aduser -searchbase "ou=2022,ou=ms,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\MS\Student Email Addresses\2022.csv" -NoTypeInformation
get-aduser -searchbase "ou=2023,ou=ms,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\MS\Student Email Addresses\2023.csv" -NoTypeInformation
get-aduser -searchbase "ou=2024,ou=ms,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\MS\Student Email Addresses\2024.csv" -NoTypeInformation

get-aduser -searchbase "ou=2025,ou=ues,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\UES\Student Email Addresses\2025.csv" -NoTypeInformation
get-aduser -searchbase "ou=2026,ou=ues,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\UES\Student Email Addresses\2026.csv" -NoTypeInformation
get-aduser -searchbase "ou=2027,ou=ues,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\UES\Student Email Addresses\2027.csv" -NoTypeInformation

get-aduser -searchbase "ou=2028,ou=les,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\LES\Student Email Addresses\2028.csv" -NoTypeInformation
get-aduser -searchbase "ou=2029,ou=les,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\LES\Student Email Addresses\2029.csv" -NoTypeInformation
get-aduser -searchbase "ou=2030,ou=les,ou=students,ou=users,ou=domain objects,dc=domain,dc=org" -filter {Enabled -eq "true" -and lastLogonTimeStamp -like "*"} -prop Surname, GivenName, SamAccountName, EmailAddress | Select Surname,Givenname,SamAccountName,EmailAddress | sort-object SamAccountName | export-csv "\\domain.org\Shares\TeacherShare\LES\Student Email Addresses\2030.csv" -NoTypeInformation