<#
 .Synopsis
  Генератор простых отчетов

 .Description
  Генератор простых отчетов

 .Parameter Count 
 [По-умолчанию: 3000] количество выбираемых событий
 .Parameter Makves_url
 URL-адрес сервера Makves. Например: http://127.0.0.1:8000/events/explore
 .Parameter Makves_user
 Имя пользователя Makves под которым данные отправляются на сервер
 .Parameter Makves_pwd
 Пароль пользователя Makves под которым данные отправляются на сервер

 .Parameter Start
 Метка времени для измения файлов в формате "yyyyMMddHHmmss"
 .Parameter Startfn
 Имя файла для метки времени

 .Example
   # Пример запуска без выделения текста
   Test-EventLog -Folder "c:\\work\\test" -Outfilename folder_test

 .Example
   # Сбор всех типов событий с компьютера dc.acme.local
   Test-EventLog -Folder -Computers dc.acme.local

 .Example
   Сбор всех типов событий (Logon/Logon) с компьютера dc.acme.local 
   Test-EventLog -Folder -Computers dc.acme.local -Target Logon

#>
function New-MakvesSimpleReport {
    Param(
        [string]$makves_url = "http://127.0.0.1:8000/events/explore",
        [string]$makves_user = "admin",
        [string]$makves_pwd = "admin",
        [string]$templatefile = "",
        [string]$outfilename = "",
        [string]$title = "",
        [string]$path = "",
        $mail = $null,
        [scriptblock]$preprocessor = $null
    )
    Import-Module Poshstache -Verbose
    $pair = "${makves_user}:${makves_pwd}"
    
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    
    $basicAuthValue = "Basic $base64"
    
    $headers = @{ Authorization = $basicAuthValue}
    
    Add-Type -AssemblyName 'System.Net.Http'
    
    function getdata { 
        param (
            [string]$makves_url = "http://127.0.0.1:8000/events/explore",
            [string]$makves_user = "admin",
            [string]$makves_pwd = "admin"
        )
        
        Add-Type -AssemblyName 'System.Net.Http'
        
        Try
        {
            $pair = "${makves_user}:${makves_pwd}"
            $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
            $base64 = [System.Convert]::ToBase64String($bytes)
            $basicAuthValue = "Basic $base64"
            [psobject]$headers = @{ Authorization = $basicAuthValue}
    
            $response = Invoke-WebRequest -Uri $makves_url -Method Get -Headers $headers
            $jsonObj = ConvertFrom-Json $([String]::new($response.Content))
            return $jsonObj
        }
        Catch {
            Write-Host "Error send data to server:" + $PSItem.Exception.Message
            return $Null
        }
    }
    if ("" -ne $path) {
        if ("" -ne $search) {
            $path = $path + "?q=" + [System.Web.HTTPUtility]::UrlEncode($search)
        }
        $makves_url = $makves_url + $path
    }
    
    $data = getdata $makves_url $makves_user $makves_pwd
    
    if ($null -eq $data) {
        return
    }

    function convertSize (
        $bytes, 
        [int]$precision
        ) 
    {
        foreach ($i in ("bytes","KB","MB","GB","TB")) {
            if (($bytes -lt 1000) -or ($i -eq "TB")) {
                $bytes = ($bytes).ToString("F0$precision")
                return $bytes + " $i"
            } else {
                $bytes /= 1024
            }
        }
    }

    if ("items" -in $data.PSobject.Properties.Name) {

        $data.items | ForEach-Object {
            if ("basic_score" -in $_.PSobject.Properties.Name) {
                $score = [math]::Ceiling($_.basic_score * 100)
                $_ | Add-Member -MemberType NoteProperty -Name score -Value $score -Force

                if ($_.basic_score -ge 0.8) {
                    $_ | Add-Member -MemberType NoteProperty -Name score_color -Value "#f86c6b" -Force
                } else {
                    if ($_.basic_score -ge 0.3) {
                        $_ | Add-Member -MemberType NoteProperty -Name score_color -Value "#ffc107" -Force
                    } else {
                        $_ | Add-Member -MemberType NoteProperty -Name score_color -Value "#4dbd74" -Force
                    }
                }
            }

            if ("size" -in $_.PSobject.Properties.Name) {
                $size_string = convertSize $_.size 2
                $_ | Add-Member -MemberType NoteProperty -Name size_string -Value $size_string -Force            
            }
        }
    }
    if ("data" -in $data.PSobject.Properties.Name) {
        $d =  $data.data
        if ("basic_score" -in $d.PSobject.Properties.Name) {
            $score = [math]::Ceiling($d.basic_score * 100)
            $d | Add-Member -MemberType NoteProperty -Name score -Value $score -Force

            if ($d.basic_score -ge 0.8) {
                $d | Add-Member -MemberType NoteProperty -Name score_color -Value "#f86c6b" -Force
            } else {
                if ($d.basic_score -ge 0.3) {
                    $d | Add-Member -MemberType NoteProperty -Name score_color -Value "#ffc107" -Force
                } else {
                    $d | Add-Member -MemberType NoteProperty -Name score_color -Value "#4dbd74" -Force
                }
            }
        }

        if ("size" -in $d.PSobject.Properties.Name) {
            $size_string = convertSize $d.size 2
            $d | Add-Member -MemberType NoteProperty -Name size_string -Value $size_string -Force            
        }
    }
    
    $data | Add-Member -MemberType NoteProperty -Name title -Value  $title -Force


    $reportTime = get-date -f "dd.MM.yyyy hh:mm"
    $data | Add-Member -MemberType NoteProperty -Name report_time -Value  $reportTime -Force
    




    if ($null -ne $preprocessor) {
        $preprocessor.Invoke($data)
    }
    
    $JSON = $data | ConvertTo-Json -Depth 5
    
    $jsonString = $JSON.ToString()
    Import-Module Poshstache
    $res = ConvertTo-PoshstacheTemplate -InputFile $templatefile -ParametersObject $jsonString 
    $res | Out-File $outfilename -Force -Encoding "UTF8"

    

    if (($null -ne $mail) -and ("" -ne $mail)) {

        $SMTPServer = $mail.server
        $SMTPClient = New-Object Net.Mail.SMTPClient($SmtpServer, $mail.port)
        $SMTPClient.EnableSSL = $mail.EnableSSL
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($mail.user, $mail.pwd);
    
        # From Core @ stackoverflow.com
        $emailMessage = New-Object System.Net.Mail.MailMessage
        $emailMessage.From = $mail.from
        foreach ($recipient in $mail.to)
        {
            $emailMessage.To.Add($recipient)
        }
        $emailMessage.IsBodyHtml = $true
        $emailMessage.Subject = $title
        $emailMessage.Body = $res
        # Do we have any attachments?
        # If yes, then add them, if not, do nothing
        ##if ($Arry_EmailAttachments.Count -ne $NULL)
        ##{
        ##    $emailMessage.Attachments.Add()
        ##}
        $SMTPClient.Send($emailMessage)
    
    }
    
}

Export-ModuleMember -Function New-MakvesSimpleReport