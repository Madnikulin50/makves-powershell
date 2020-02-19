[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] $computers = (""),
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $outfilename = "dc_logs",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $Count = 1000,
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $Exclude = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $user = "current",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $pwd = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $start = "",
    [string]$startfn = "", ##".dc-log-monitor.time_mark",
    [string]$makves_url = "", ##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
)

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

if (($startfn -ne "") -and (Test-Path $startfn)) {
   Try
   {
       $start = Get-Content $fnstart
   }
   Catch {
       Write-Host "Error read time mark:" + $PSItem.Exception.Message
       $start = ""
   }
} 


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
  
function store($data) {
    if ($outname -ne "") {
        $data | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
    }
    
    if ($uri -ne "") {
        $data | Add-Member -MemberType NoteProperty -Name Forwarder -Value "event-forwarder" -Force
        $JSON = $data | ConvertTo-Json
        Try {
            $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
            Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
            Write-Host "Send data to server:" + $data.Name
        }
        Catch {
            Write-Host "Error send data to server:" +  $PSItem.Exception.Message
        }
    }
}


function execute () {
    Foreach ($computer in $computers)
    {
        Write-host "Checking Computer $Computer"
        $FilterHashProperties = @{Logname='security'; ID=4776,4672,4624,4634,4800,4801}
        if ($start -ne "") {
            Write-Host "start: " $start
            $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
            $FilterHashProperties.Add("startTime", $starttime)
        }

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
            ##$Events = $Events | Select-Object -first $Count
            $excludeFilter = 'SYSTEM|LOCAL|NETWORK|СИСТЕМА|.*\$$'
            if ($exclude -ne "") {
                $excludeFilter = $excludeFilter + '|' + $exclude
            }
            $counter = 0
            
            $Events = $Events | Foreach-Object {
                $cur = $_ 
                if ($cur.Properties[1].Value -notmatch $excludeFilter) {
                    try {
                        $xml = $_.ToXml()
                        $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                    } Catch {
                        Write-Host "error create xml" -ForegroundColor Red
                    }
                    store($cur)
                    $counter ++
                }
                

            }
            Write-host "Found at least $($counter) events ! Here are the $NumberOfLastEventsToGet last ones"
                
        }
        Catch {
            $msg = "Error accessing Event Logs of $computer by Get-WinEvent + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
        }
    }
}

execute


if ($startfn -ne "") {
    $markTime | Out-File -FilePath $startfn -Encoding UTF8
    Write-Host "Store new mark: " $markTime
}
