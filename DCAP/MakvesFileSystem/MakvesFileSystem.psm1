<#
 .Synopsis
  Сбор данных о файлах в папке

 .Description
  Сбор данных о файлах в папке

 .Parameter Folder
 [Необязательный] Корневая папка(локальная или сетевая) для сбора данных
 
 .Parameter Base
 [Необязательный] Корневая OU для зачитывания списка компьтеров из ActiveDirectory
 .Parameter Server 
 [Необязательный] Имя домен-контроллера для зачитывания списка компьтеров из ActiveDirectory
 .Parameter User
 [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
 .Parameter Pwd
 [Необязательный] пароль пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
 .Parameter Outfilename
 Имя файла результатов
 .Parameter extruct
 Выделять текст из doc, docx, xls, xlsx
 .Parameter Compliance
 Проверять файлы на соответсвие текста стандартам
 .Parameter No_hash 
 Не производить вычисление хэша файлов
 .Parameter Makves_url
 URL-адрес сервера Makves. Например: http://192.168.0.77:8000
 .Parameter Makves_user
 Имя пользователя Makves под которым данные отправляются на сервер
 .Parameter Makves_pwd
 Пароль пользователя Makves под которым данные отправляются на сервер
 .Parameter Start
 Метка времени для измения файлов в формате "yyyyMMddHHmmss"
 .Parameter Startfn
 Имя файла для метки времени
 .Parameter Logfn
 Имя файла для лога


 .Example
   # Пример запуска без выделения текста
   Test-FS -Folder "c:\\work\\test" -Outfilename folder_test

 .Example
   # Пример запуска c выделениeм текста
   Test-FS -Folder "c:\\work\\test" -Outfilename folder_test -Extruct

 .Example
   # Пример запуска без выделения текста сбора данных о всех папках общего доступа, с компьютеров зарегистрированных в указанной организационной единице
   Test-FS -Folder -Base "DC=acme,DC=local -Server "dc.acme.local" -Outfilename "folder_test"

 .Example
   Пример запуска с проверкой текста на соответствие стандартам
   Test-FS -Folder "c:\\work\\test" -Outfilename folder_test -Compliance

#>
function Test-FileSystem {
    param (
        [string]$folder = "C:\work",
        [string]$outfilename = "folder", ##"",
        [string]$base = "",
        [string]$server = "",
        [int]$hashlen = 1048576,
        [switch]$no_hash = $false,
        [switch]$extruct = $false,
        [switch]$compliance = $true,
        [switch]$monitor = $false,
        [string]$start = "",
        [string]$startfn = "", ##".file-monitor.time_mark",
        [string]$makves_url =  "", ##"http://192.168.2.22:8000",
        [string]$makves_user = "admin",
        [string]$makves_pwd = "admin",
        [string]$logfilename = ""##""
    )

    ## Init web server 
    $uri = $makves_url + "/data/upload/file-info"
    $pair = "${makves_user}:${makves_pwd}"

    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)

    $basicAuthValue = "Basic $base64"

    $headers = @{ Authorization = $basicAuthValue}

    if ($makves_url -eq "") {
        $uri = ""
        Add-Type -AssemblyName 'System.Net.Http'
    }
    $scriptFolder = $MyInvocation.MyCommand.Module.ModuleBase
    Write-Host "Script folder:" + $scriptFolder


    if ($compliance -eq $true) {
        Import-Module $scriptFolder"/compliance.dll" -Verbose
    }


    

    $markTime = Get-Date -format "yyyyMMddHHmmss"

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



    $LogDate = get-date -f yyyyMMddhhmm 
    $outfile = ""

    if ($outfilename -ne "") {
        $outfile = "$($outfilename)_$LogDate.json"
        if (Test-Path $outfile) 
        {
            Remove-Item $outfile
        }
    }

    $logfile = ""

    if ($logfilename -ne "") {
        $logfile = "$($logfilename)_$LogDate.json"
        if (Test-Path $logfile) 
        {
            Remove-Item $logfile
        }
    }

    Write-Host "base: " $folder
    Write-Host "outfile: " $outfile
    Write-Host "outfile: " $logfilename



    Function Get-MKVS-FileHash([String] $FileName,$HashName = "SHA1") 
    {
        if ($hashlen -eq 0) {
            $FileStream = New-Object System.IO.FileStream($FileName,"Open", "Read") 
            $StringBuilder = New-Object System.Text.StringBuilder 
            [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($FileStream)| ForEach-Object {[Void]$StringBuilder.Append($_.ToString("x2"))} 
            $FileStream.Close() 
            $StringBuilder.ToString()
        } else {
            $StringBuilder = New-Object System.Text.StringBuilder 
            $binaryReader = New-Object System.IO.BinaryReader(New-Object System.IO.FileStream($FileName,"Open", "Read"))
        
            $bytes = $binaryReader.ReadBytes($hashlen)
            $binaryReader.Close() 
            if ($bytes -ne 0) {
                [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($bytes)| ForEach-Object { [Void]$StringBuilder.Append($_.ToString("x2")) }
            }
            $StringBuilder.ToString()
        }
    }

    function Get-MKVS-DocText([String] $FileName) {
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Word.DisplayAlerts = 0
        $text = ""
        Try
        {
            $catch = $false
            Try{
                $Document = $Word.Documents.Open($FileName, $null, $null, $null, "")
            }
            Catch {
                Write-Host 'Doc is password protected.`r`n'
                if ($logfilename -ne "") {
                    'Doc is password protected.`r`n' | Out-File -FilePath $logfile -Encoding UTF8 -Append
                }
                $catch = $true
            }
            if ($catch -eq $false) {
                $Document.Paragraphs | ForEach-Object {
                    $text += $_.Range.Text
                }
                
            }
        }
        Catch {
            Write-Host $PSItem.Exception.Message
            $Document.Close()
            $Word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
            Remove-Variable Word
        }
        $Document.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
        Remove-Variable Word        
        return $text
    }

    function Get-MKVS-XlsText([String] $FileName) {
        $excel = New-Object -ComObject excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = 0
        $text = ""
        $password
        Try    
        {
            $catch = $false
            Try{
                $wb =$excel.Workbooks.open($path, 0, 0, 5, "")
            }
            Catch{
                Write-Host 'Book is password protected.'
                if ($logfilename -ne "") {
                    'Book is password protected.`r`n' | Out-File -FilePath $logfile -Encoding UTF8 -Append
                }
                $catch = $true
            }
            if ($catch -eq $false) {
                foreach ($sh in $wb.Worksheets) {
                    #Write-Host "sheet: " $sh.Name            
                    $endRow = $sh.UsedRange.SpecialCells(11).Row
                    $endCol = $sh.UsedRange.SpecialCells(11).Column
                    #Write-Host "dim: " $endRow $endCol
                    for ($r = 1; $r -le $endRow; $r++) {
                        for ($c = 1; $c -le $endCol; $c++) {
                            $t = $sh.Cells($r, $c).Text
                            $text += $t
                            #Write-Host "text cel: " $r $c $t
                        }
                    }
                }
            }
        }
        Catch {
            Write-Host $PSItem.Exception.Message
            if ($logfilename -ne "") {
                "Extruct Excel text: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
        }
        #Write-Host "text: " $text
        $excel.Workbooks.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        Remove-Variable excel
        return $text
    }

    function Get-MKVS-FileText([String] $FileName, [String] $Extension) {
        #Write-Host "filename: " $FileName
        #Write-Host "ext: " $Extension

        switch ($Extension) {
            ".doc" {
                return Get-MKVS-DocText $FileName
            }
            ".docx" {
                return Get-MKVS-DocText $FileName
            }
            ".xls" {
                return Get-MKVS-XlsText $FileName
            }
            ".xlsx" {
                return Get-MKVS-XlsText $FileName
            }
        }
        return ""    
    }

    function inspectFile($cur) {
        $cur = $cur | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
        Write-Host $cur.FullName
        if ($logfilename -ne "") {
            "Start inspect file $($cur.FullName)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
        }
        Try
        {
            $acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
        }
        Catch {
            Write-Host "Get-Acl error:" + $PSItem.Exception.Message
            if ($logfilename -ne "") {
                "Get-Acl error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
        }
        $path = $cur.FullName
        $ext = $cur.Extension
            
        if ($cur.PSIsContainer -eq $false) {
            if ($no_hash -eq $false) {
            Try
            {
                $hash = Get-MKVS-FileHash $path
                }
                Catch {
                    if ($logfilename -ne "") {
                        "Get-MKVS-FileHash error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                    Write-Host $PSItem.Exception.Message
                    Try
                    {
                        $hash = Get-FileHash $path | Select-Object -Property "Hash"
                    }
                    Catch {
                            Write-Host $PSItem.Exception.Message
                            if ($logfilename -ne "") {
                                "Get-FileHash error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                            }
                    }
                }
            }

            if ($extruct -eq $true)
            {
                Try
                {
                    $text =  Get-MKVS-FileText $path $ext
                    $cur | Add-Member -MemberType NoteProperty -Name Text -Value $text -Force
                }
                Catch {
                    Write-Host "Get-MKVS-FileText error:" + $PSItem.Exception.Message
                    if ($logfilename -ne "") {
                        "Get-MKVS-FileText error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                }    
            }

            if ($compliance -eq $true)
            {
                Try
                {
                    $compliance =  Get-Compliance -File $path
                    
                    $cur | Add-Member -MemberType NoteProperty -Name Compliance -Value $compliance -Force
                }
                Catch {
                    Write-Host "Get-Compliance error:" + $PSItem.Exception.Message
                    if ($logfilename -ne "") {
                        "Get-Compliance error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                }

                ##Try
                ##{
                ##    $content =  Get-Document-Text -File $path
                ##    Write-Host "Get-Compliance error:" + $PSItem.Exception.Message
                ##    $cur | Add-Member -MemberType NoteProperty -Name Content -Value $content -Force
                ##}
                ##Catch {
                ##	Write-Host "Get-Document-Text error:" + $PSItem.Exception.Message
                ##}
            }

            $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
        }
            
        $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
        $cur | Add-Member -MemberType NoteProperty -Name Computer -Value $env:computername -Force
        Try
        {
            if ($outfile -ne "") {
                $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
            }
            
            if ($uri -ne "") {
                $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "folder-forwarder" -Force
                $JSON = $cur | ConvertTo-Json
                Try
                {
                    $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());

                    Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
                    Write-Host  "Send data to server:" + $cur.Name
                    if ($logfilename -ne "") {
                        "Send data to server: $($cur.Name)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                }
                Catch {
                    Write-Host "Error send data to server:" + $PSItem.Exception.Message
                    if ($logfilename -ne "") {
                        "Error send data to server: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                }
            }
        }
        Catch {
            Write-Host "Store error:" + $PSItem.Exception.Message
            if ($logfilename -ne "") {
                "Store error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
        }
    }

    function inspectFolder($f) {
        Try
        {
            $cur  = Get-Item $f  
        } Catch {
            Write-Host "Error Get-Item:" + $f + ":" $PSItem.Exception.Message
            if ($logfilename -ne "") {
                "Error Get-Item $($f): $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
            return;
        }

        if ($cur -eq $null) {
            return;
        }

        $cur  = $cur | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
        
        Write-Host $cur.FullName
        Try
        {
            $acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"  
        } Catch {
            Write-Host "Error Get-Acl:" + $f + ":" $PSItem.Exception.Message
            if ($logfilename -ne "") {
                "Error Get-Acl $($f): $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
            return;
        }
        $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
        $cur | Add-Member -MemberType NoteProperty -Name RootAudit -Value $true -Force
        $cur | Add-Member -MemberType NoteProperty -Name Computer -Value $env:computername -Force
    
        if ($outfile -ne "") {
            $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
        }
        
        if ($uri -ne "") {
            $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "folder-forwarder" -Force
            $JSON = $cur | ConvertTo-Json
            Try
            {
                $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
                Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
                Write-Host  "Send data to server:" + $cur.Name
                if ($logfilename -ne "") {
                    "Send data to server: $($cur.Name)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                }
            }
            Catch {
                Write-Host "Error send data to server:" + $PSItem.Exception.Message
                if ($logfilename -ne "") {
                    "Error send data to server: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                }
            }
        }

        
        if ($start -ne "") {
            Write-Host "start: " $start
            if ($logfilename -ne "") {
                "start:  $($start)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
            }
            $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)

            Get-ChildItem $f -Recurse | Where-Object { $_.LastWriteTime -gt $starttime } | Foreach-Object {
                Try
                {
                    inspectFile $_
                }
                Catch {
                    if ($logfilename -ne "") {
                        "inspectFile error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                    Write-Host "inspectFile error:" + $PSItem.Exception.Message
                }
            }
        } else {
            Get-ChildItem $f -Recurse | Foreach-Object {
                Try
                {
                    inspectFile $_
                }
                Catch {
                    Write-Host "inspectFile error:" + $PSItem.Exception.Message
                    if ($logfilename -ne "") {
                        "inspectFile error: $($PSItem.Exception.Message)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
                    }
                }
            }
        }
    }


    if ($base -eq "" ) {
        inspectFolder $folder
    } else {
        Import-Module ActiveDirectory
        $GetAdminact = Get-Credential
        $computers = Get-ADComputer -Filter * -server $server -Credential $GetAdminact -searchbase $base | Select-Object "Name"    
        $computers | ForEach-Object {
            $machine = $_.Name
            Write-Host "export shares from machine: " $machine
            net view $machine | Select-Object -Skip  7 | Select-Object -SkipLast 2|
            ForEach-Object -Process {[regex]::replace($_.trim(),'\s+',' ')} |
            ConvertFrom-Csv -delimiter ' ' -Header 'sharename', 'type', 'usedas', 'comment' |
            foreach-object {
                inspectFolder "\\$($machine)\$($_.sharename)"
            }
        }

    }

    if ($startfn -ne "") {
        $markTime | Out-File -FilePath $startfn -Encoding UTF8
        Write-Host "Store new mark: " $markTime
        if ($logfilename -ne "") {
            "Store new mark: $($markTime)`r`n" | Out-File -FilePath $logfile -Encoding UTF8 -Append
        }
    }

    if ($monitor -eq $true) {
        $filter = '*.*'  # You can enter a wildcard filter here. 


        $fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $true; NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 
        
        # Here, all three events are registerd.  You need only subscribe to events that you need: 
        
        Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
            $name = $Event.SourceEventArgs.Name 
            $changeType = $Event.SourceEventArgs.ChangeType 
            $timeStamp = $Event.TimeGenerated 
            Write-Host "The file '$name' was $changeType at $timeStamp" -fore green 
            $fullname = Join-Path -Path $folder -ChildPath $name
            inspectFile $fullname
        } 
        
        Register-ObjectEvent $fsw Deleted -SourceIdentifier FileDeleted -Action { 
            $name = $Event.SourceEventArgs.Name 
            $changeType = $Event.SourceEventArgs.ChangeType 
            $timeStamp = $Event.TimeGenerated 
            Write-Host "The file '$name' was $changeType at $timeStamp" -fore red 
        } 
        
        Register-ObjectEvent $fsw Changed -SourceIdentifier FileChanged -Action { 
            $name = $Event.SourceEventArgs.Name 
            $changeType = $Event.SourceEventArgs.ChangeType 
            $timeStamp = $Event.TimeGenerated 
            Write-Host "The file '$name' was $changeType at $timeStamp" -fore white
            $fullname = Join-Path -Path $folder -ChildPath $name
            inspectFile $fullname
        } 




        while ($true)
        {
            Start-Sleep -Milliseconds 1000
            if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character))
            {
                Write-Host "You pressed CTRL-C. Do you want to continue doing this and that?" 
                $key = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
                if ($key.Character -eq "N") { break; }
            }
        }

        Unregister-Event FileDeleted 
        Unregister-Event FileCreated 
        Unregister-Event FileChanged
    }
}

Export-ModuleMember -Function Test-FileSystem