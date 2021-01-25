#Send Message to remote computer (probably will only work in local Zone)

[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')


$title = 'Computer Name'
$msg   = 'Enter the Computer Name to Send Message To:'
$compName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

$title = 'Message'
$msg   = 'Enter the Message You want to send:'
$message = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

Invoke-WmiMethod -Path Win32_process -Name Create -ArgumentList "msg * $message" -Computername $compName

