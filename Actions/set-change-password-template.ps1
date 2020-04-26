
#region header
param (
    [string]$base = "",
    [string]$admin_user = "current",
    [string]$admin_pwd = ""
)
if ($user -eq "current") {
    $GetAdminact = $null 
}
else {
    if ($user -ne "") {
        $pass = ConvertTo-SecureString -AsPlainText $pwd -Force    
        $GetAdminact = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pass    
    }
    else {
        $GetAdminact = Get-Credential
    }
}
#endregion

#region item
Set-ADUser -Identity {{.NTName}} -ChangePasswordAtLogon $true
#endregion

#region footer
#end region



