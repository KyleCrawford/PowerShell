
##################################################################################################
#Program Details:
# Used to retreive Computer Asset number, SerialNumber and Hard drive SerialNumber, model info and saves this info to a 
# .CSV file located in a network share folder
#Author:
# Kyle Crawford
#Notes
# Requires to be run from a physical location on the computer, either from the hard drive or a usb in the computer
# Requires admin access to ensure accurate results from the hard drive serial query
# Does not handle a computer that has multiple hard drives
#Last Updated:
# 25/9/2017
##################################################################################################


#if not in admin, open a new script as admin
#doesn't work in the console (needs a running environment)

#*note does not work if run from a network location
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
    #No Administrative rights, it will display a popup window asking user for Admin rights
    
    #retreives running location (requires to be run as a script)
    $arguments = "& " + $myinvocation.mycommand.definition + ""

    #attempt to open elevated script
    Start-Process "$psHome\powershell.exe" -Verb runAs -ArgumentList $arguments 

    #close the old, Unelevated script
    break
}

#the regular expression to test against the user's input
$regEx = '.*-[0-9][0-9][0-9][0-9]-N-[0-9][0-9][0-9]'

#prompt the user for the deployment ID, and loop back if the information is not in the correct format.
do {
    $depID = (read-host "enter the deployment ID in the format '[anything]-1234-N-123'").ToUpper() 
}
#check if user has entered in the ID in the correct format, if not, loop back and ask again.
until ( $depID -match $regEx )


######## To edit the save path ################


#the path to save the computer info to
$date = Get-Date -UFormat "%Y.%m.%d"

                                                                                      # eg. JSG-1234-N
$savePath = "\\edm-goa-smsd-22\GOA_Refresh\Deployment_Team_Documents\ScriptManifests\" + $depID.substring(0,$depID.length - 4)

#Check to see if there is a folder with the same month and deployment ID
if (-not (test-path $savePath))
{
    #if it doesn't, create the folder
    New-Item $savePath -type directory
}

#add the trimmed deployment ID to the savePath and append .csv
$savePath += "\$date.csv"

#eg of final savePath: 
#\\edm-goa-smsd-22\GOA_Refresh\Deployment_Team_Documents\ScriptManifests\JSG-1234-N\2017.12.30

########### End of SavePath Edit ############################


#Check to see if there is a CSV file already, if so, Import it to a variable
if ((test-path $savePath) -eq $true)
{
    #Import CSV file to be able to append additional info
    $oldData = [Object[]] (Import-Csv $savePath)
}

do
{
    #prompt user to see if computer is going to surplus, storage, or keep on site
    $allocation = Read-host "Computer going to (S)urplus (D)MP storage or (K)eep on site"
}
until ($allocation.ToLower() -match "^s" -or $allocation.ToLower() -match "^d" -or $allocation.ToLower() -match "^k")

#switch statement to allocate the correct information to the variable to allow for similar input on all forms
switch -Wildcard ($allocation.ToLower())
{
    "s*" { $allocation = "Surplus" }
    "d*" { $allocation = "DMP Storage" }
    "k*" { $allocation = "Keep on Site" }
}

#the system serial number
$serialNumber = (wmic bios get serialnumber)[2]

#the make and model of the system
$makeModel = (Get-WMIobject -Class Win32_ComputerSystem | select-object -property Model).model
            
#the main hard drive of the system (does not handle multiple hard drive disks)
$HDDSerial = ((get-wmiobject -class Win32_physicalMedia | where-object {$_.tag -like "*DRIVE0"} | Select -property Serialnumber).SerialNumber).Trim()

#Adds Deployment ID, device allocation, Asset tag, system serial number, HDD serial number and computer make/model to the $newData variable
$newData = New-Object PSObject -Property @{ "Deployment ID" = $depID; Allocation = $allocation;  Asset = hostname; SerialNumber = $serialNumber; `
    "Make/Model" = $makeModel; HDDSerial = $HDDSerial; Quantity = 1;}

#do-until loop to keep trying to export the file if the file is already in use
do
{
    #Tries to export the info to a .csv file, if the file is in use, will prompt the user to either try again, or to quit without saving the info to the file
    try
    {
        #Selects the properties to ensure the correct order when adding them to the .csv file
        $oldData + $newData | select-object -property "Deployment ID", Allocation, Asset, SerialNumber, HDDSerial, Make/Model, Quantity | Export-Csv $savePath -NoTypeInformation

        #running sum of all the items in the list. 
        #(insert '= sum(e1:e#) at the bottom)
        
        #if the above line is successful, changes $input var to not be 'x' to allow script to exit loop
        $input = "x"
    }

    #catches the event of the file being open or unavailable. 
    catch
    {
        #Prompts user to enter 'c' to try again, or 'x' to quit without saving
        $input = read-host "File in use, enter 'c' to try again, or 'x' to exit without saving"
    }
} until ($input -eq "x")


