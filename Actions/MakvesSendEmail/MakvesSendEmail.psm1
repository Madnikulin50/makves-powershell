<#
 .Synopsis
  Отправка информации по почте

 .Description
  Отправка информации по почте

 .Parameter From
 [Необязательный] Отправитель
 
 .Parameter To
 Получатель

 .Parameter Server 
 [Необязательный] Почтовый сервер

 .Parameter Port 
 [Необязательный] Порт почтового сервера

 .Parameter UseSsl
 Использовать Ssl

 .Parameter Subject
 Тема письма

 .Parameter Body
 Тело письма
 
 .Example
   # Пример запуска без выделения текста
   Send-Email -From 'User01 <user01@fabrikam.com>' -To 'ITGroup <itdept@fabrikam.com>' -Subject "Don't forget today's meeting!" -Body "Some Body" -UseSsl

#>
function Send-Email {
    param (
        [string]$from = "admin@makves.ru",
        [string[]]$to = (""),
        [string]$server = $PSEmailServer,
        [int]$port = 25,
        [bool]$usessl = $false,
        [string]$subject = "",
        [string]$body = ""
    )

    Write-Host "From:" + $from
    Write-Host "To:" + $to
    Write-Host "Subject:" + $subject

    if ($usessl -eq $true) {
        Send-MailMessage -From $from -To $to -Subject $subject -Body $body `
        -SmtpServer $server -Port $port -UseSsl
    } else {
        Send-MailMessage -From $from -To $to -Subject $subject -Body $body `
        -SmtpServer $server -Port $port
    }
}
