param (
    [string]$EmailFrom = '', ##Var of sender email
    [string]$EmailTo = '{{.Email}}', ##var of User mail
    [string]$Subject = 'Администратор: Требуется объяснить.', ##Subject is text
    [string]$Body = '{{.BODY}}', ##Body of mail
    [string]$SmtpServer = '',##server smtp like "smtp.makves.ru"
    [string]$userName = '{{.NTName}}' ## User name from Active Ditrectory
     )
    if ($Body -eq ''){
        $body = 'Добрый день! Уважаемый ' +$userName + ' Ваша учетная запись имеет доступ к ... Прошу Вас дать пояснение.'
    }

$smtp = New-Object net.mail.smtpclient($SmtpServer)
$smtp.Send($EmailFrom, $EmailTo, $Subject, $Body)

