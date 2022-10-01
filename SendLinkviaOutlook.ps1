#File has to be in %appdata%\Microsoft\Windows\SendTo


param($p1)
$Outlook = New-Object -ComObject Outlook.Application

$Mail = $Outlook.CreateItem(0)
$Mail.Subject = ""

$Mail.HTMLBody = "<HTML><BODY>Hier der Link zum Dokument: <a href=""" + $PSBoundParameters["p1"] + """>" + $PSBoundParameters["p1"] + "</a>.</BODY></HTML>"

$Mail.Display(0)