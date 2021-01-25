[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')



$title = 'Computer Name'
$msg   = 'Enter a Computer Name:'

$PCName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

if ($PCName -ne "")
{
   Start-Process chrome "https://$($PCName):4343 -nowindow" 
}


