param (
    [string]$base = "",
    [string]$server = "",
    [string]$outfilename = 'export_ad',
    [string]$user = "current",
    [string]$pwd = "",
    [switch]$notping = $true,
    [string]$start = "",
    [string]$startfn = "", ##".ad-monitor.time_mark",
    [string]$makves_url = "",##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [int]$timeout = 0,
    [switch]$restrict_fileds = $true,
    [string[]]$fields = ("Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
    "sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
    "PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "ObjectClass", "DNSHostName", "ObjectCategory", "UserPrincipalName", "ServicePrincipalName",
    "GivenName", "Surname", "sn", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
    "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
    "Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
    "Mail", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
    "CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
    "Manager", "logonCount", "LogonHours", "employeeNumber")
 )

Write-Host "base: " $base
Write-Host "server: " $server

Write-Host "user: " $user
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesActiveDirectory"

Import-Module -Name $scriptPath"\MakvesActiveDirectory" -Verbose

Test-ActiveDirectory -Base $base -Server $Server -Outfilename $outfilename `
 -User $user -Pwd $pwd `
 -NotPing $notping `
 -Start $start -StartFn $startfn `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd `
 -Restrict_Fields $restrict_fileds -Fields $fields

Write-Host "Export finished..."  -ForegroundColor Green
