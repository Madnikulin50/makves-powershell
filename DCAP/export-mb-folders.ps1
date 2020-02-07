param (
    [string]$outfilename = ".\explore-mailboxes",
    [string]$makves_url = "", ##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"    
 )
 
 ## Init web server 
 $uri = $makves_url + "/agent/push"
 $pair = "${makves_user}:${makves_pwd}"
 
 $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
 $base64 = [System.Convert]::ToBase64String($bytes)
 
 $basicAuthValue = "Basic $base64"
 
 $headers = @{ Authorization = $basicAuthValue}
 
 if ($makves_url -eq "") {
     $uri = ""
     Add-Type -AssemblyName 'System.Net.Http'
 }
 
 
 if ($compliance -eq $true) {
     Import-Module "./compliance.dll" -Verbose
 }
 
 $markTime = Get-Date -format "yyyyMMddHHmmss"
 
  if ($startfn -ne "") {
     Try
     {
         $start = Get-Content $startfn
     }
     Catch {
         Write-Host "Error read time mark:" + $PSItem.Exception.Message
         $start = ""
     }
 } 
 
 
 
 $LogDate = get-date -f yyyyMMddhhmm 
 $outfile = ""
 
 if ($outfilename -ne "") {
     $outfile = "$($outfilename)_$LogDate.json"
     if (Test-Path $outfile) 
     {
         Remove-Item $outfile
     }
 }
 
 Write-Host "outfile: " $outfile

 function inspect($item) {
  $t = $item 
  $t | Add-Member -MemberType NoteProperty -Name Type -Value "exchange-mailbox" -Force        
  $t | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-mailboxes-forwarder" -Force
  


  $JSON = $t | ConvertTo-Json
  Try
  {
      if ($outfile -ne "") {
          $JSON | Out-File -FilePath $outfile -Encoding UTF8 -Append
      }
     
      if ($uri -ne "") {
          Try
          {
              $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());

              $resp = Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
              #Write-Host  "Send data to server:" + $cur.Name
          }
          Catch {
              Write-Host "Error send data to server:" + $PSItem.Exception.Message
          }
      }
  }
  Catch {
      Write-Host $PSItem.Exception.Message
  }
}


function execute() {
    Get-Mailbox -resultSize Unlimited | Select-Object Alias, Name, DisplayName, ServerName |
    foreach-object {
        try
        {
            $cur = $_
            try{
                $fstat = Get-MailboxFolderStatistics -Identity $cur.alias
                $cur | Add-Member -MemberType NoteProperty -Name Folders -Value $fstat -Force
            }
            Catch {
                $msg = "No Get-MailboxFolderStatistics: + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Yellow
            }
            

            try{
                $s = Get-MailboxStatistics -Identity $cur.alias | Select-Object DisplayName, LastLogonTime, ItemCount, LastLogoffTime, LegacyDN, LastLoggedOnUserAccount, ObjectClass
                #$s | ForEach-Object {
                #    $cs = $_ | ConvertTo-Json
                # Write-Host "Statistic: " $cs
                #}
                $cur | Add-Member -MemberType NoteProperty -Name Statistic -Value $s -Force
            }
            Catch {
                $msg = "No Get-MailboxStatistics: + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Yellow
            }

            
            try{
                $uc= Get-MailboxUserConfiguration -Identity ($cur.alias)
                if ($null -ne $uc) {
                    #Write-Host "UserConfiguration: " $uc
                    $cur | Add-Member -MemberType NoteProperty -Name UserConfiguration -Value $uc -Force
                }
            }
            Catch {
                $msg = "No Get-MailboxUserConfiguration: + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Yellow
            }
            
            try {
                $p = Get-MailboxPermission -Identity ($cur.alias) | Select-Object Identity, User, AccessRights
                if ($null -ne $p) {
                    [string]$t = "["
                    $p | foreach-object {
                        if ($t.length -ne 1) {
                            $t += ","
                        }
                        $t += "{`"Identity`": `"$($_.Identity)`","
                        $t += "`"User`": `"$($_.User)`","
                        $t += "`"AccessRights`": `"$($_.AccessRights)`"}"
                    }
                    $t += "]"
                    #Write-Host "permision: " $t
                    $cur | Add-Member -MemberType NoteProperty -Name Permissions -Value $t -Force
                }
            } Catch {
                $msg = "No Get-MailboxPermission: + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Yellow
            }

            try {
                $p = Get-MailboxFolderPermission -Identity ($cur.alias + ":\") -ErrorAction SilentlyContinue | Select-Object Identity, FolderName, User, AccessRights
                if ($null -ne $p) {
                    [string]$t = "["
                    $p| foreach-object {
                        if ($t.length -ne 1) {
                            $t += ","
                        }
                        $t += "{`"Identity`": `"$($_.Identity)`","
                        $t += "`"FolderName`": `"$($_.FolderName)`","
                        $t += "`"User`": `"$($_.User)`","
                        $t += "`"AccessRights`": `"$($_.AccessRights)`"}"
                    }
                    $t += "]"
                    #Write-Host "folder permision: " $t
                    $cur | Add-Member -MemberType NoteProperty -Name FolderPermissions -Value $t -Force
                }
            } Catch {
                $msg = "No Get-MailboxFolderPermission: + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Yellow
            }

            $msg = "Writing " + $cur.Alias + "-" + $cur.Name

            Write-Host $msg
            
            inspect($cur)
    }
    Catch {
                $msg = "Common + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Yellow
    }
    }
    if ($startfn -ne "") {
        $markTime | Out-File -FilePath $startfn -Encoding UTF8
        Write-Host "Store new mark: " $markTime
    }
}

Write-Host $PSItem.Exception.Message -ForegroundColor Green
