param (
    [string]$base = 'DC=otr,DC=ru',
    [string]$server = '172.31.4.5',
    [string]$outfilename = 'export_ad',
    [string]$user = "",
    [string]$pwd = "",
    [switch]$force = $false,
    [string]$start = "",
    [string]$makves_url = "",##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
 )

Write-Host "base: " $base
Write-Host "server: " $server

Write-Host "user: " $user
Write-Host "pwd: " $pwd
#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm

Import-Module ActiveDirectory

$SearchBase = $base 

## Init web server 
$uri = $makves_url + "/data/upload/agent"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}


$outfile = ""

if ($outfilename -ne "") {
    $outfile = "$($outfilename)_$LogDate.json"
    if (Test-Path $outfile) 
    {
        Remove-Item $outfile
    }
}

Write-Host "outfile: " $outfile

$domain = Get-ADDomain -server $server

Write-Host "domain: " $domain.NetBIOSName

if ($outfile -ne "") {
  $domain | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}
if ($uri -ne "") { 
  $domain | Add-Member -MemberType NoteProperty -Name Forwarder -Value "ad-forwarder" -Force
  $JSON = $domain | ConvertTo-Json
  $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
  Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
}



if ($start -ne "") {
  Write-Host "start: " $start
  $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
}

function Get-ADPrincipalGroupMembershipRecursive() {
  Param(
      [string] $dsn,
      [array]$groups = @()
  )

  $obj = Get-ADObject -server $server $dsn -Properties memberOf

  foreach( $groupDsn in $obj.memberOf ) {

      $tmpGrp = Get-ADObject -server $server  $groupDsn -Properties * | Select-Object "Name", "cn", "distinguishedName", "objectSid", "DisplayName", "memberOf"

      if( ($groups | Where-Object { $_.DistinguishedName -eq $groupDsn }).Count -eq 0 ) {
          $add = $tmpGrp 
          $groups +=  $tmpGrp           
          $groups = Get-ADPrincipalGroupMembershipRecursive $groupDsn $groups
      }
  }

  return $groups
}


Get-ADUser -server $server -searchbase $SearchBase -Filter * | Foreach-Object {
$cur = $_
try {
$filter = 'Name -eq "' + $_.Name + '"'
Write-Host $filter
$cur = Get-ADUser -server $server -searchbase $SearchBase -Filter $filter -Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} | Select-Object "Name", "GivenName", "Surname", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
"CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
"Manager", "Enabled", "lastlogondate", "ObjectClass", "logonCount", "LogonHours", "UserPrincipalName", "ServicePrincipalName"
$cur = $cur[0]
} catch {
    Write-Host $error 
}
    


  if ($start -ne "") {
    if (($cur.whenChanged -lt $starttime) -and ($cur.lastlogondate -lt $starttime)){
      Write-Host "skip " $cur.Name
      return
    }
    Write-Host "write " $cur.Name

  }

  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"

  if ($cur.thumbnailPhoto -ne $null) {
    $cur.thumbnailPhoto =[Convert]::ToBase64String($cur.thumbnailPhoto)
  }

  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force

  $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  if ($outfile -ne "") {
    $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
  }
  if ($uri -ne "") { 
    $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "ad-forwarder" -Force
    $JSON = $cur | ConvertTo-Json
    $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
    Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
  }
}

Write-Host "Export finished..."
