<#
 .Synopsis
  Сбор событий EventLog c удаленных рабочих станций

 .Description
  Сбор событий EventLog c удаленных рабочих станций

 .Parameter Computers
 Список компьютеров с которых необходимо собрать логи событий
 
 .Parameter Target
 Типы собираемых событий
 
 .Parameter Outfilename
 Имя файла результатов

 .Parameter User
 [Необязательный] Имя пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
 .Parameter Pwd
 [Необязательный] пароль пользователя под которым производится запрос. Если не заданно, то выводится диалог с запросом |
 

 .Parameter Count 
 [По-умолчанию: 3000] количество выбираемых событий
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
function Test-EventLog {
    Param(
    [string[]]$computers = (""),
    [string]$outfilename = "events",
    [int32]$Count = 3000,
    [string]$user = "",
    [string]$pwd = "",
    [string]$start = "",
    [string] $fwd = "",
    [ValidateSet("All","Logon","Service","User","Computer", "Clean", "File", "MSSQL", "RAS", "USB", "Printer", "Sysmon", "TS", "Policy")] [string[]]$target=("All"),
    [string]$startfn = "", ##".event-monitor.time_mark",
    [string]$makves_url = "", ##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$exclude_user = "",
    [bool]$split_by_id = $false
)

    $NumberOfLastEventsToGet = $Count


    ## Init web server 
    $uri = $makves_url + "/data/upload/event"
    $pair = "${makves_user}:${makves_pwd}"
    
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    
    $basicAuthValue = "Basic $base64"
    
    $headers = @{ Authorization = $basicAuthValue}
    
    if ($makves_url -eq "") {
        $uri = ""
        Add-Type -AssemblyName 'System.Net.Http'
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
    
    Write-Host "computers: " $computers
    Write-Host "outfile: " $outfile
    Write-Host "target: " $target
    
    
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
    
    
    $DebugPreference = "Continue"
    $ErrorActionPreference = "SilentlyContinue"
    
    $FilterHashProperties = $null
    
    
    function store($data) {
        if ($outname -ne "") {
            $data | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
        }
       
        if ($uri -ne "") {
            $data | Add-Member -MemberType NoteProperty -Name Forwarder -Value "event-forwarder" -Force
            $JSON = $data | ConvertTo-Json
            Try
            {
                $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
                $resp = Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
            }
            Catch {
                Write-Host "Error send data to server:" +  $PSItem.Exception.Message
            }
        }
    }
    
    
    
    
    function IsEmpty($Param){
        If ($Param -eq "All" -or $Param -eq "" -or $null -eq $Param -or $Param -eq 0) {
            Return $True
        } Else {
            Return $False
        }
    }
    
    
    Write-Host "Starting script..."
    function ExportFor($eid, $ln, $type, $sp) {
        if ($true -eq $sp) {
            $eid | ForEach-Object {
                ExportFor $_ $ln $type $false
            }
            return
        }
    
        if ($fwd -ne "") {
            $ln = $fwd
        }        
    
        Write-Host "logname:" $ln
        Write-Host "type: " $type
       
        
    
        $FilterHashProperties = @{
            LogName = $ln
        }
        
        if ($start -ne "") {
            Write-Host "start: " $start
            $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
            $FilterHashProperties.Add("startTime", $starttime)
        }
    
        If (!(IsEmpty $eid)){
            Write-Host "eid: " $eid
            $FilterHashProperties.Add("ID",$eid)
        }
     
          
        $msg = ("About to collect events on $($computers.count)") + $(If ($($computers.count) -gt 1){" machines"}Else{" machine"})
        Write-host $msg
        
        Foreach ($computer in $computers)
        {
            $msg = "Checking Computer $Computer"
            Write-host $msg -BackgroundColor yellow -ForegroundColor Blue
            
            try
            {
                if ($null -eq $GetAdminact) {
                    if ("" -eq $Computer) {
                        $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -ErrorAction SilentlyContinue -MaxEvents $Count
                    } else {
                        $Events = Get-WinEvent -FilterHashtable $FilterHashProperties -Computer $Computer -ErrorAction SilentlyContinue -MaxEvents $Count
                    }
                } else {
                    if ("" -eq $Computer) {
                        $Events = Get-WinEvent -Credential $GetAdminact -FilterHashtable $FilterHashProperties -ErrorAction SilentlyContinue -MaxEvents $Count
                    } else {
                        $Events = Get-WinEvent -Credential $GetAdminact -FilterHashtable $FilterHashProperties -Computer $Computer -ErrorAction SilentlyContinue -MaxEvents $Count
                    }
                }
                $res = $Events | Foreach-Object {
                    $cur = $_ 
                    try {
                        $xml = $_.ToXml()
                        $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                    } Catch {
                        Write-Host "error create xml" -ForegroundColor Red
                    }
                    store($cur)
    
                }
                Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones"     
            }
            Catch {
                $msg = "Error accessing Event Logs of $computer by Get-WinEvent + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Red
                try {    
                    $Events = get-eventlog -logname $ln -newest 10000 -Computer $Computer
                    $Events | Where-Object {$eid -contains $_.EventID}
                    $Events | Select-Object -first $NumberOfLastEventsToGet
                    $Events | Foreach-Object {
                        $cur = $_ 
                        try {
                            $xml = $_.ToXml()
                            $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                        } Catch {
                            Write-Host "error create xml" -ForegroundColor Red
                        }
                        
                        $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                        store($cur)
                    }
                    
                }
                Catch {
                    $msg = "Error accessing Event Logs of $computer by get-eventlog + $PSItem.Exception.InnerExceptionMessage"
                    Write-Host $msg -ForegroundColor Red
                }
    
            }
            Finally {
                Write-Host "OK_"
            }
        }    
    }
    
    
    function execute() {
        Foreach ($i in $target)
        {
            if ($i -eq "Logon" -or $i -eq "All") {
                ExportFor ("4776","4672", "4624", "4634", "4800", "4801") "Security" "logon"  $split_by_id
            }
    
            if ($i -eq "Service" -or $i -eq "All") {
                ExportFor ("7036","7031") "System" "service"  $split_by_id
            }
    
            if ($i -eq "User" -or $i -eq "All") {
                ExportFor ("4720", "4722", "4723", "4724", "4725", "4726", "4728", "4729", "4732", "4735", "4738", "4740", "4741", "4742", "4743", "4767", "4780", "4794", "5376", "5377", "4755", "4756", "4757", "5135", "5136", "5137", "5138", "5139", "5141", "4761", "4760") "Security" "user"  $split_by_id
            }
    
            if ($i -eq "Computer" -or $i -eq "All") {
                ExportFor ("4721", "4720", "4722", "4725", "4726", "4728", "4729", "4738", "4740", "4741", "4742", "4743", "4767") "Security" "computer"  $split_by_id
            }
    
            if ($i -eq "Clean" -or $i -eq "All") {
                
                Write-Host "EventID: " $id_clean
                ExportFor ("1102") "Security" "clean"  $split_by_id
            }
    
            if ($i -eq "File" -or $i -eq "All") {
                ExportFor("5140", "5142", "5143", "5144", "5145") "Security" "file"
                ExportFor("4656", "4663", "4660", "4670", "4658") "Security" "file"
            }
            if ($i -eq "Printer" -or $i -eq "All") {
                ExportFor ("307")  ("Microsoft-Windows-PrintService/Operational") "printer"  $split_by_id
            }
    
            if ($i -eq "MSSQL" -or $i -eq "All") {
                ExportFor ("18456")  "Application" "mssql"  $split_by_id
            }
    
            if ($i -eq "RAS" -or $i -eq "All") {
                ExportFor ("20249", "20250", "20253", "20255", "20258", "20266", "20271", "20272") "RemoteAccess/Operational" "ras"  $split_by_id
            }
    
            if ($i -eq "USB" -or $i -eq "All") {
                ExportFor ("2003") "Microsoft-Windows-DriverFrameworks-UserMode/Operational" "usb"  $split_by_id
            }
            if ($i -eq "Sysmon" -or $i -eq "All") {
                ExportFor ("1", "3", "5", "11", "12", "13", "14") "Microsoft-Windows-Sysmon/Operational" "sysmon"  $split_by_id
            }    
            if ($i -eq "TS" -or $i -eq "All") {
                ExportFor ("21", "24") "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" "ts"  $split_by_id
            }

        }
    
        Write-Host "Iteration done."
    }
    
    execute
    
    
    
    if ($startfn -ne "") {
        $markTime | Out-File -FilePath $startfn -Encoding UTF8
        Write-Host "Store new mark: " $markTime
    }
    
    Write-Host "Events export done" -ForegroundColor Green
}

Export-ModuleMember -Function Test-EventLog