param (
    [string]$base = 'DC=acme,DC=loc',
    [string]$server = '172.18.5.10',
    [string]$outfilename = 'export_ldap',
    [string]$user = "f@masters",
    [string]$pwd = "f",
    [switch]$notping = $true,
    [string]$start = "",
    [string]$startfn = ".ad-monitor.time_mark_172.18.5.10",
    [string]$makves_url = "",##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [int]$timeout = 0
 )
 
Write-Host "base: " $base
Write-Host "server: " $server

Write-Host "user: " $user
Write-Host "pwd: " $pwd

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

if (($startfn -ne "") -and (Test-Path $startfn))  {
  Try
  {
      $start = Get-Content $startfn
  }
  Catch {
      Write-Host "Error read time mark:" + $PSItem.Exception.Message
      $start = ""
  }
} 


Write-Host "outfile: " $outfile
if ($user -eq "current") {
  $GetAdminact = $null 
} else {
  if ($user -ne "") {
      $pass = ConvertTo-SecureString -AsPlainText $pwd -Force    
      $GetAdminact = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pass    
  } else {
      $GetAdminact = Get-Credential
  }
}

function store($cur) {
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

if ($start -ne "") {
  Write-Host "start: " $start
  $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
}

function inspectComputer($cur) {
  if ($start -ne "") {
    if (($cur.whenChanged -lt $starttime) -and ($cur.lastlogondate -lt $starttime)) {
      Write-Host "skip " $cur.Name
      return
    }
    Write-Host "write " $cur.Name

  }

  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
  if ($notping -eq $false) {
    if ($null -eq $GetAdminact) {
      $licensies = Get-WmiObject SoftwareLicensingProduct -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Description, LicenseStatus
    } else {
      $licensies = Get-WmiObject SoftwareLicensingProduct -Credential $GetAdminact -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Description, LicenseStatus
    }
    if ($licensies -ne $Null) {
      Write-Host $cur.DNSHostName " : " $($licensies)
      $cur | Add-Member -MemberType NoteProperty -Name OperatingSystemLicensies -Value $licensies -Force
    }
  
    Try {
      if ($null -eq $GetAdminact) {
        $userprofiles = Get-WmiObject  -Class win32_userprofile -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object sid, localpath
      } else {
        $userprofiles = Get-WmiObject -Credential $GetAdminact -Class win32_userprofile -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object sid, localpath
      }
      if ($userprofiles -ne $null) {
        Write-Host $cur.DNSHostName  " : " $userprofiles
        $cur | Add-Member -MemberType NoteProperty -Name UserProfiles -Value $userprofiles -Force
      }    
    } Catch {
      Write-Host $cur.DNSHostName  " : " "$($_.Exception.Message)"
    }
  
    Try {
      if ($null -eq $GetAdminact) {
        $apps = Get-WMIObject -Class win32_product -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Name, Version
      } else {
        $apps = Get-WMIObject -Class win32_product -Credential $GetAdminact -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Name, Version
      }
      if ($apps -ne $Null) {
        Write-Host $cur.DNSHostName " : " $apps
        $cur | Add-Member -MemberType NoteProperty -Name Applications -Value $apps -Force
      }
      
    }
    Catch {
        Write-Host $cur.DNSHostName " win32_product Offline "
        try {
         
  
          $Registry = $Null;
          Try{$Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $cur.DNSHostName);}
          Catch{Write-Host -ForegroundColor Red "$($_.Exception.Message)";}
          
          If ($Registry){
            $apps =  New-Object System.Collections.Generic.List[System.Object];
            $UninstallKeys = $Null;
            $SubKey = $Null;
            $UninstallKeys = $Registry.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall",$False);
            $UninstallKeys.GetSubKeyNames()|%{
              $SubKey = $UninstallKeys.OpenSubKey($_,$False);
              $DisplayName = $SubKey.GetValue("DisplayName");
              If ($DisplayName.Length -gt 0){
                $Entry = $Base | Select-Object *
                $Entry.ComputerName = $ComputerName;
                $Entry.Name = $DisplayName.Trim(); 
                $Entry.Publisher = $SubKey.GetValue("Publisher"); 
                [ref]$ParsedInstallDate = Get-Date
                If ([DateTime]::TryParseExact($SubKey.GetValue("InstallDate"),"yyyyMMdd",$Null,[System.Globalization.DateTimeStyles]::None,$ParsedInstallDate)){					
                $Entry.InstallDate = $ParsedInstallDate.Value
                }
                $Entry.EstimatedSize = [Math]::Round($SubKey.GetValue("EstimatedSize")/1KB,1);
                $Entry.Version = $SubKey.GetValue("DisplayVersion");
                [Void]$apps.Add($Entry);
              }
            }
            
              If ([IntPtr]::Size -eq 8){
                      $UninstallKeysWow6432Node = $Null;
                      $SubKeyWow6432Node = $Null;
                      $UninstallKeysWow6432Node = $Registry.OpenSubKey("Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",$False);
                          If ($UninstallKeysWow6432Node) {
                              $UninstallKeysWow6432Node.GetSubKeyNames()|%{
                              $SubKeyWow6432Node = $UninstallKeysWow6432Node.OpenSubKey($_,$False);
                              $DisplayName = $SubKeyWow6432Node.GetValue("DisplayName");
                              If ($DisplayName.Length -gt 0){
                                $Entry = $Base | Select-Object *
                                  $Entry.ComputerName = $ComputerName;
                                  $Entry.Name = $DisplayName.Trim(); 
                                  $Entry.Publisher = $SubKeyWow6432Node.GetValue("Publisher"); 
                                  [ref]$ParsedInstallDate = Get-Date
                                  If ([DateTime]::TryParseExact($SubKeyWow6432Node.GetValue("InstallDate"),"yyyyMMdd",$Null,[System.Globalization.DateTimeStyles]::None,$ParsedInstallDate)){                     
                                  $Entry.InstallDate = $ParsedInstallDate.Value
                                  }
                                  $Entry.EstimatedSize = [Math]::Round($SubKeyWow6432Node.GetValue("EstimatedSize")/1KB,1);
                                  $Entry.Version = $SubKeyWow6432Node.GetValue("DisplayVersion");
                                  $Entry.Wow6432Node = $True;
                                  [Void]$apps.Add($Entry);
                                }
                              }
                        }
                      }
             Write-Host $cur.DNSHostName + " : " $apps
             $cur | Add-Member -MemberType NoteProperty -Name Applications -Value $apps -Force
          }
        } Catch {
          Write-Host $cur.DNSHostName " error apps" "$($_.Exception.Message)"
        }
     }
  
  }
  

  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  store($cur)
}

function inspectGroup($cur) {
  if ($start -ne "") {
    if ($cur.whenChanged -lt $starttime) {
      Write-Host "skip " $cur.Name
      return
    }

  }

  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
  
   $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  store($cur)
}


function inspectUser($cur) {
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

  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  store($cur)
}



function execute() {
    $dn = New-Object System.DirectoryServices.DirectoryEntry ("LDAP://$($server):389/$base",$user,$pwd)

    # Here look for a user
    $ds = new-object System.DirectoryServices.DirectorySearcher($dn)
    $ds.filter = "(|(&(objectCategory=person)(objectClass=user))(objectCategory=computer)(objectCategory=group))"
    $ds.SearchScope = "subtree"
    $ds.PropertiesToLoad.Add("distinguishedName")
    $ds.PropertiesToLoad.Add("sAMAccountName")
    $ds.PropertiesToLoad.Add("lastLogon")
    $ds.PropertiesToLoad.Add("telephoneNumber")
    $ds.PropertiesToLoad.Add("memberOf")
    $ds.PropertiesToLoad.Add("distinguishedname")
    $ds.PropertiesToLoad.Add("otherHomePhone");
    $ds.PropertiesToLoad.Add("Name");
    $ds.PropertiesToLoad.Add("sn");
    $ds.PropertiesToLoad.Add("dn");
    $ds.PropertiesToLoad.Add("whenCreated");
    $ds.PropertiesToLoad.Add("memberOf");
    $ds.PropertiesToLoad.Add("badPwdCount");
    $ds.PropertiesToLoad.Add("objectSid");
    $ds.PropertiesToLoad.Add("DisplayName");
    $ds.PropertiesToLoad.Add("IPv4Address");
    $ds.PropertiesToLoad.Add("IPv6Address");
    $ds.PropertiesToLoad.Add("OperatingSystem");
    $ds.PropertiesToLoad.Add("OperatingSystemHotfix");
    $ds.PropertiesToLoad.Add("OperatingSystemServicePack");
    $ds.PropertiesToLoad.Add("OperatingSystemVersion");
    $ds.PropertiesToLoad.Add("PrimaryGroup");
    $ds.PropertiesToLoad.Add("ManagedBy");
    $ds.PropertiesToLoad.Add("userAccountControl");
    $ds.PropertiesToLoad.Add("Enabled");
    $ds.PropertiesToLoad.Add("lastlogondate");
    $ds.PropertiesToLoad.Add("ObjectClass");
    $ds.PropertiesToLoad.Add("DNSHostName");
    $ds.PropertiesToLoad.Add("ObjectCategory");
    $ds.PropertiesToLoad.Add("LastBadPasswordAttempt");
    $ds.PropertiesToLoad.Add("ServicePrincipalName");
    $ds.PropertiesToLoad.Add("GivenName");
    $ds.PropertiesToLoad.Add("Surname");
    $ds.PropertiesToLoad.Add("StreetAddress");
    $ds.PropertiesToLoad.Add("City");
    $ds.PropertiesToLoad.Add("state");
    $ds.PropertiesToLoad.Add("PostalCode");
    $ds.PropertiesToLoad.Add("Country");
    $ds.PropertiesToLoad.Add("Title");
    $ds.PropertiesToLoad.Add("Company");
    $ds.PropertiesToLoad.Add("Description");
    $ds.PropertiesToLoad.Add("Department");
    $ds.PropertiesToLoad.Add("OfficeName");
    $ds.PropertiesToLoad.Add("thumbnailPhoto");
    $ds.PropertiesToLoad.Add("Mail");
    $ds.PropertiesToLoad.Add("Manager");
    $ds.PropertiesToLoad.Add("logonCount");
    $ds.PropertiesToLoad.Add("UserPrincipalName");
    $ds.PropertiesToLoad.Add("ServicePrincipalName");
    $ds.PropertiesToLoad.Add("ObjectClass");
    $ds.PropertiesToLoad.Add("cn");
    $ds.PropertiesToLoad.Add("whenChanged");
    $ds.PropertiesToLoad.Add("badPwdCount");
    $ds.PropertiesToLoad.Add("DisplayName");
    $ds.PropertiesToLoad.Add("Enabled");
    $ds.PropertiesToLoad.Add("telephoneNumber");
    $ds.PropertiesToLoad.Add("PasswordNeverExpires");
    $ds.PropertiesToLoad.Add("PasswordExpired");
    $ds.PropertiesToLoad.Add("DoesNotRequirePreAuth");
    $ds.PropertiesToLoad.Add("CannotChangePassword");
    $ds.PropertiesToLoad.Add("PasswordNotRequired");
    $ds.PropertiesToLoad.Add("TrustedForDelegation");
    $ds.PropertiesToLoad.Add("TrustedToAuthForDelegation")
    $ds.PropertiesToLoad.Add("LogonHours");
    $ds.PropertiesToLoad.Add("employeeNumber");
 
    


    $ds.FindAll() | foreach-object {
        $cur = $_.Properties
        store $cur
    }
}


Write-Host "Export finished..."