
#########################################################################################################################
# Description:
# Used for deployments, this script copies user data (_LOCALdata, Downloads, Documents, Favourites Chrome bookmarks),
# Collects information for network drives and network printers which are then added on the new computer
# 
# *Notes*
# This script is designed to be used at the computer while the user is logged in as it takes information from the user that is currently logged in.
# Retrieval of user's network drives and printers takes place in the non-elevated portion of the script since when the script is elevated
# it will copy the information from the admin user's account and not the currently logged in user.


## Auto Deploy (At user station)

#Parameters 
#$userName - the user name that the files are to be transferred to (needed since when the script is run as admin it wants to use admin username.
#$saveRoot - the drive that the script is being run from (needed since when it is run as admin it uses local hard drive (usually C:\))
#$oldOrNew - Used as input from the user, stores either 'old' or 'new' representing if we are on the old or new computer.

Param(
  #  [Parameter(Mandatory=$False, Position=1)]
    [string]$userName,
  #  [Parameter(Mandatory=$False, Position=2)]
    [string]$saveRoot,
  #  [Parameter(Mandatory=$False, Position=3)]
    [string]$oldOrNew
    )

#Opens the CD tray, prompts the user, then closes it after user presses enter at prompt
Function Check-CDTray
{
    #creates a new object that represents the CD tray
    $Diskmaster = New-Object -ComObject IMAPI2.MsftDiscMaster2 
    $DiskRecorder = New-Object -ComObject IMAPI2.MsftDiscRecorder2 
    $DiskRecorder.InitializeDiscRecorder($DiskMaster) 

    #Eject the CD tray
    $DiskRecorder.EjectMedia() 
    #Prompt user to inform them that the CD tray has opened
    (New-Object -ComObject wscript.shell).popup("Opening CD tray... Click 'OK' to continue", 0, "Don't Forget!", 1)
    #Close the CD tray
    $DiskRecorder.CloseTray()
}


# if the script is not running with admin rights
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    #prompt user to see if they are on the old or new computer
    #loop until user has entered in one of the correct entries
    #User can enter anything starting with 'n' or 'o'
    do
    {
        $runWhich = read-host "Are you on the (o)ld or (n)ew computer?"
    }
    until ($runWhich.ToLower() -match "^o" -or $runWhich.ToLower() -match "^n")

    #assign the appropriate value back to $runwhich, to allow for a sanitized input.
    switch -Wildcard ($runWhich.ToLower())
    {
        "o*" {$runWhich = "old"}
        "n*" {$runWhich = "new"}
    }

    #Filepath of script to be run (this script)
    $arguments = "& " + $myinvocation.mycommand.definition

    #add the user to args to pass to elevated script
    #need to pass current user since when you run the script as admin it uses the admin username
    $arguments += " $env:USERNAME"

    #getting the drive that this script is being run from.
    #used since when the script is run as admin it uses the local hard drive (usually C:\)
    $path = (Get-Location).drive.Root
    $arguments += " $path"

    #passing either 'old' or 'new' representing which computer we are on currently.
    $arguments += " $runWhich"

    #launch the script as administrator passing the above args
    Start-Process "$psHome\powershell.exe" -Verb runAs -ArgumentList $arguments
    
    #Adding the filepath of where the data is stored
    $path = $path + "\UserData\$env:USERNAME"

    #if we are running on the old computer
    if ($runWhich -eq "old")
    {
        #create the user folder on the USB where everything will be stored
        #check to see if the folder exists, if it doesn't, create the folder
        if (-not (test-path $path))
        {
            #create new folder using the $path
            New-Item -Path $path -ItemType "directory"
        }

        Write-host "Retrieving network printer info..."
        #Get list of all the network printers that are connected to this computer and save it to a .csv file
        Get-WmiObject Win32_Printer | select -property Name | where-object {$_.Name -like "\\*"} | export-csv -Path ($path + "\printers.csv")

        Write-host "Retrieving network drives info..."
        #Get list of all the network drives that are connected to this computer and save it to a .csv file
        Get-WmiObject -Class Win32_MappedLogicalDisk | Select Name, ProviderName | export-csv -Path ($path + "\Drives.csv") 
    } 
    
    #If we are running on the new computer
     elseif ($runWhich -eq "new")
     {
        # Install network printers
        #get csv file from file stored on usb (requires this script to have been run on the old computer)
        $csv = Import-Csv -Path ($path + "\printers.csv")

        #Add each network printer that is in the .csv file
        foreach ($item in $csv)
        {
            (new-Object -ComObject WScript.Network).AddWindowsPrinterConnection($item.name)
        }
        write-host "Done adding printers"


        # Install network drives
        $drives = Import-Csv -Path ($path + "\Drives.csv")

        #add each network drive that is in the .csv file
        foreach ($d in $drives)
        {
            #$d.Name.substring(0,$d.name.length - 1)

            #removes the ':' from the drive letter
            $driveLetter = $d.Name.substring(0,$d.name.length - 1)
            #adds the network drive to the computer
            New-PSDrive –Name $driveLetter –PSProvider FileSystem –Root $d.ProviderName –Persist
        }

        Write-Host "Done adding drives"
     }

     #Else statement to catch errors in $runWhich var
     else
     {
        #inform user of error. User must press enter to continue.
        Read-host "Something went wrong. Press enter to end script."
     }

    #Closes the unelevated script
    break
}


#the filepath to save the user data to
$savePath = "$saveRoot" + "UserData\$userName"

#the base path to copy from
$copyPath = "c:\Users\$userName"

#if we are on the old computer
if ($oldOrNew -eq "old")
{
    #create the folder on the USB to save everything:
    #check to see if the folder exists, if it doesn't, create the folder
    if (-not (test-path $savePath))
    {
        #create new folder using the $path
        New-Item -Path $savePath -ItemType "directory"
    }
     
      
# Copy files to USB
    #copy my docs
    robocopy "$copyPath\Documents" "$savePath\Documents" /e

    #copy Downloads
    robocopy "$copyPath\Downloads" "$savePath\Downloads" /e

    #copy _localDATA
    robocopy "c:\_localdata" "$savePath\_Localdata" /e

    robocopy "$copyPath\Favorites" "$savePath\Favorites" /e

    #copy Chrome Bookmarks
    robocopy "C:\Users\$userName\AppData\Local\Google\Chrome\User Data\Default" $savePath "Bookmarks" 

    #(to be added: firefox bookmarks)

    # Open CD tray
    Check-CDTray

    read-host -Prompt "Check for shared mailboxes" 
}

elseif ($oldOrNew -eq "new")
{
    #read-host "we are new"

    # Transfer files from USB to computer
    #open chrome to prepare folder stucture for bookmarks to be transferred back
    #(done first since it may take a while to open)

    #temp comment out to test line below
    Invoke-Item "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    #start chrome.exe "https://sharedservices.gov.ab.ca/pmg/pjs/odtt"

    #the file path that the data is to be transferred from
    $from = $savePath

    #The base file path to move the files to
    $to = "c:\Users\$userName"

    #copy over Downloads
    robocopy "$from\Downloads" "$to\Downloads" /e /move

    #copy over Documents
    robocopy "$from\Documents" "$to\Documents" /e /move

    #copy over Favorites
    robocopy "$from\Favorites" "$to\Favorites" /e /move

    #Copy over _LocalData
    robocopy "$from\_Localdata" "C:\_Localdata" /e /move

    #prompt user to wait until chrome has opened, then continue
    read-host "Wait for Chrome to open, then press enter."

    #Copy Chrome Bookmarks (requires chrome to have been launched once first)
    robocopy "$from" "C:\Users\$userName\AppData\Local\Google\Chrome\User Data\Default" "Bookmarks" /move

    #prompt user to check if the files have been successfully transferred
    read-host "Check for successful transfer. Leftover user files will be deleted after enter is pressed."

    #checks for any leftover files and deletes them to ensure no user data is left on the USB
    if (test-path -Path "$from\Downloads"){
        remove-item -recurse -path "$from\Downloads" -Force }
    
    if (test-path -Path "$from\Documents"){
        remove-item -recurse -path "$from\Documents" -Force }

    if (test-path -Path "$from\Favorites"){
        remove-item -recurse -path "$from\Favorites" -Force }

    if (test-path -Path "$from\_Localdata"){
        remove-item -recurse -path "$from\_Localdata" -Force }

    if (test-path -Path "$from\Bookmarks"){
        remove-item -recurse -path "$from\Bookmarks" -Force }

    #Prompts the user to see if they want to delete the user folder from the USB drive.
    do {
        $delete = Read-Host "Ok to delete user folder from USB? Y/N"
    }until ($delete.ToLower() -match "^y" -or $delete.ToLower() -match "^n")

    #If the user said yes to deleting the folder, delete the folder
    if ($delete.ToLower() -match "^y")
    {
        if (test-path -Path $from){
            remove-item -recurse -path $from -Force } 
    }
}

#else statement here to catch any unexpected errors
else
{
    read-host "Something went wrong, try again. Press enter to end script."
    break
}




read-host "End of script. Press enter to continue."

