$serverpath = "E:\TeacherData\UES"
foreach ($user in Get-ChildItem $serverpath) {
    $homedir = $serverpath+"\"+$user
    # Try to see if we can read the ACL (if we can, then we can continue, if not, we need to take ownership
    try {
        takeown /R /D Y /F $homedir
        icacls $homedir /reset /t /c /q
        icacls $homedir /setowner "example\administrator" /t /c /q
        # Once this is ran, see if it was successful by getting the ACL List. If not, something else is messed up, so exit
        try {
            $acl = Get-Acl $homedir
            }
        catch {
            write-host Failed to take ownership
            sys.exit
        }
    }
    Catch {
        write-host Failed to take ownership
        sys.exit
    }
$rights=[System.Security.AccessControl.FileSystemRights]::FullControl
$inheritance=[System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
$propagation=[System.Security.AccessControl.PropagationFlags]::InheritOnly
$allowdeny=[System.Security.AccessControl.AccessControlType]::Allow
$dirace=New-Object System.Security.AccessControl.FileSystemAccessRule ($user,$rights,$inheritance,$propagation,$allowdeny)
$ACL.AddAccessRule($dirace)
$rights=[System.Security.AccessControl.FileSystemRights]::FullControl
$inheritance=[System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
$propagation=[System.Security.AccessControl.PropagationFlags]::None
$allowdeny=[System.Security.AccessControl.AccessControlType]::Allow
$dirace=New-Object System.Security.AccessControl.FileSystemAccessRule ($user,$rights,$inheritance,$propagation,$allowdeny)
$ACL.AddAccessRule($dirace)
Set-Acl $homedir $acl
    # Now do the same thing for every subfolder
    Get-ChildItem $homedir | foreach {        
        icacls $homedir"\"$_ /t /c /q /reset
        icacls $homedir"\"$_ /setowner $user /t /c /q
    }
    # Reset Ownership back to the person's whose home directory this is
    icacls $homedir /setowner $user /t /c /q
}
