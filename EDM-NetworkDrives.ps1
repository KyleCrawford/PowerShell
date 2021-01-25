[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

#To allow user to double click to open Powershell script, Right click script > Open With > choose another app > More apps > look for another app on this PC > 
#Powershell .exe located - C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
#all powershell scripts can now be ran by double clicking

#edm information plus current logged in user
$userDrive = "\\networkPath" + $env:username

#prompt user for what letter to add as
$title = 'Shared Drive Letter'
$msg   = 'Enter the Desired Drive Letter (eg. H):'

$loopCount = 0
$breakLoop = $false


#loop allows the user to make mistakes several times. If the user does not enter a drive letter twice, the script will close (This includes hitting cancel or the 'x')
Do
{
    $driveLetter = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    if ($driveLetter.Length -eq 0 -and $loopCount -gt 0)
    {
        [System.Windows.MessageBox]::Show("No Drive Letter Entered. Ending Script")
        exit
    }

    elseif ($driveLetter.Length -eq 0)
    {
        [System.Windows.MessageBox]::Show("Must enter Drive Letter")
        $loopCount++
    }
    elseif ($driveLetter.Length -gt 1)
    {
        [System.Windows.MessageBox]::Show("Drive Letter must be only one character. IE. 'H'")
    }
    else
    {
        #we got 1 (cannot be less than 0)
        $breakLoop = $true
    }

} While (!$breakLoop)

#map personal drive
New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $userDrive -Persist

#run login script for shared network drives
CSCRIPT \\NetworkPath\EDM.LogonScript01.cmd