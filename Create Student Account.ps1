Import-Module activedirectory
#Generated Form Function
function GenerateForm {
########################################################################
# Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.10.0
# Generated On: 9/5/2013 9:16 PM
# Generated By: PenglaseB
########################################################################

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$Cancel = New-Object System.Windows.Forms.Button
$Create = New-Object System.Windows.Forms.Button
$GradYear = New-Object System.Windows.Forms.ComboBox
$GradYearLabel = New-Object System.Windows.Forms.Label
$StudentID = New-Object System.Windows.Forms.TextBox
$StudentIDLabel = New-Object System.Windows.Forms.Label
$LastName = New-Object System.Windows.Forms.TextBox
$LastNameLabel = New-Object System.Windows.Forms.Label
$FirstNameLabel = New-Object System.Windows.Forms.Label
$FirstName = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$Create_OnClick= 
{
    # Grab and format some variables
    $firstname = $firstname.Text
    $firststrip = $firstname -replace " ",""
    $firststrip = $firststrip -replace ".``",""
    $firststrip = $firststrip -replace ".`'",""
    $FirstInitial = $firstname.substring(0,1)
    $lastname = $lastname.Text
    $laststrip = $lastname -replace " ",""
    $laststrip = $laststrip -replace ".``",""
    $laststrip = $laststrip -replace ".`'",""
    $oldstyle = $laststrip+$Firstinitial 
    $displayname = $lastname+", "+$firstname
    $samaccountname = $laststrip+$firststrip
    $logonname = $laststrip+$firststrip
    $id = $StudentID.Text
    # If the SAMAccountName ends up over 20, we have to shorten it for 'backwards compatibility'
    if ($samaccountname.Length -gt 20) {
        $samaccountname = $samaccountname.Substring(0,20)
    }
    $domain = "@example.org"
    $studentdomain = "@students.example.org"
    $fileserver = "\\fs1.example.org\students\"
    $gradyear = $gradyear.Text
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
                $changepass = 0}
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
            #try {
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
                Add-ADGroupMember -Identity $gradyear -Member $samaccountname
                New-Item $fileserver$gradyear"\"$samaccountname -Type Directory -Force
                $acl = Get-ACL $fileserver$gradyear"\"$samaccountname
            	$UserObj = New-Object System.Security.Principal.NTAccount($($SamAccountName))
                $rights=[System.Security.AccessControl.FileSystemRights]::FullControl
                $inheritance=[System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
                $propagation=[System.Security.AccessControl.PropagationFlags]::InheritOnly
                $allowdeny=[System.Security.AccessControl.AccessControlType]::Allow
                $dirace=New-Object System.Security.AccessControl.FileSystemAccessRule ($($samaccountname),$rights,$inheritance,$propagation,$allowdeny)
                $ACL.AddAccessRule($dirace)
                $rights=[System.Security.AccessControl.FileSystemRights]::FullControl
                $inheritance=[System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
                $propagation=[System.Security.AccessControl.PropagationFlags]::None
                $allowdeny=[System.Security.AccessControl.AccessControlType]::Allow
                $dirace=New-Object System.Security.AccessControl.FileSystemAccessRule ($($samaccountname),$rights,$inheritance,$propagation,$allowdeny)
                $ACL.AddAccessRule($dirace)                
                $acl.SetOwner($UserObj)
            	Set-ACL $fileserver$gradyear"\"$samaccountname $acl
                $smtpServer = "mail.example.org"
                $msg = new-object Net.Mail.MailMessage
                $smtp = new-object Net.Mail.SmtpClient($smtpServer)
                $msg.From = "accounts@example.org"
                $msg.To.Add("$techperson, its@example.org")
                $msg.subject = "New Account - $building"
                $msg.body = "
Creating in: $ou
Created Account for: $FirstName $LastName
Username: $samaccountname
Password: $password
email: $logonname$studentdomain"
                $smtp.Send($msg)
            #}
            #catch {
            #    write-host Error Creating User: $samaccountname
            #}
            #Finally {}
        }
    }
}

$Cancel_OnClick= 
{
#TODO: Place custom script here
$form1.close()
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 195
$System_Drawing_Size.Width = 365
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.Name = "form1"
$form1.Text = "SD Student Accounts"


$Cancel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 56
$System_Drawing_Point.Y = 150
$Cancel.Location = $System_Drawing_Point
$Cancel.Name = "Cancel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$Cancel.Size = $System_Drawing_Size
$Cancel.TabIndex = 9
$Cancel.Text = "Cancel"
$Cancel.UseVisualStyleBackColor = $True
$Cancel.add_Click($Cancel_OnClick)

$form1.Controls.Add($Cancel)


$Create.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 150
$Create.Location = $System_Drawing_Point
$Create.Name = "Create"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 125
$Create.Size = $System_Drawing_Size
$Create.TabIndex = 8
$Create.Text = "Create Account"
$Create.UseVisualStyleBackColor = $True
$Create.add_Click($Create_OnClick)

$form1.Controls.Add($Create)

$GradYear.DataBindings.DefaultDataSourceUpdateMode = 0
$GradYear.FormattingEnabled = $True
$GradYear.Items.Add("2014")|Out-Null
$GradYear.Items.Add("2015")|Out-Null
$GradYear.Items.Add("2016")|Out-Null
$GradYear.Items.Add("2017")|Out-Null
$GradYear.Items.Add("2018")|Out-Null
$GradYear.Items.Add("2019")|Out-Null
$GradYear.Items.Add("2020")|Out-Null
$GradYear.Items.Add("2021")|Out-Null
$GradYear.Items.Add("2022")|Out-Null
$GradYear.Items.Add("2023")|Out-Null
$GradYear.Items.Add("2024")|Out-Null
$GradYear.Items.Add("2025")|Out-Null
$GradYear.Items.Add("2026")|Out-Null
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 87
$GradYear.Location = $System_Drawing_Point
$GradYear.Name = "GradYear"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 25
$System_Drawing_Size.Width = 121
$GradYear.Size = $System_Drawing_Size
$GradYear.TabIndex = 7

$form1.Controls.Add($GradYear)

$GradYearLabel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 65
$GradYearLabel.Location = $System_Drawing_Point
$GradYearLabel.Name = "GradYearLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 120
$GradYearLabel.Size = $System_Drawing_Size
$GradYearLabel.TabIndex = 6
$GradYearLabel.Text = "Graduation Year"

$form1.Controls.Add($GradYearLabel)

$StudentID.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 89
$StudentID.Location = $System_Drawing_Point
$StudentID.Name = "StudentID"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$StudentID.Size = $System_Drawing_Size
$StudentID.TabIndex = 5

$form1.Controls.Add($StudentID)

$StudentIDLabel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 65
$StudentIDLabel.Location = $System_Drawing_Point
$StudentIDLabel.Name = "StudentIDLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 78
$StudentIDLabel.Size = $System_Drawing_Size
$StudentIDLabel.TabIndex = 4
$StudentIDLabel.Text = "Student ID"

$form1.Controls.Add($StudentIDLabel)

$LastName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 35
$LastName.Location = $System_Drawing_Point
$LastName.Name = "LastName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 100
$LastName.Size = $System_Drawing_Size
$LastName.TabIndex = 3

$form1.Controls.Add($LastName)

$LastNameLabel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 13
$LastNameLabel.Location = $System_Drawing_Point
$LastNameLabel.Name = "LastNameLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 19
$System_Drawing_Size.Width = 100
$LastNameLabel.Size = $System_Drawing_Size
$LastNameLabel.TabIndex = 2
$LastNameLabel.Text = "Last Name"

$form1.Controls.Add($LastNameLabel)

$FirstNameLabel.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 13
$FirstNameLabel.Location = $System_Drawing_Point
$FirstNameLabel.Name = "FirstNameLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 19
$System_Drawing_Size.Width = 78
$FirstNameLabel.Size = $System_Drawing_Size
$FirstNameLabel.TabIndex = 1
$FirstNameLabel.Text = "First Name"

$form1.Controls.Add($FirstNameLabel)

$FirstName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 35
$FirstName.Location = $System_Drawing_Point
$FirstName.Name = "FirstName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 119
$FirstName.Size = $System_Drawing_Size
$FirstName.TabIndex = 0

$form1.Controls.Add($FirstName)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm
