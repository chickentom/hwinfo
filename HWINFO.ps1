#Gets some useful hardware information out of remote PCS
#Need to Run as Administrator

#Imports Computerhostnames from the File given down below
#You need to change the path if the file is named differently
#Please make sure that there is a column named 'Computer'

$Temp = Import-Csv "./converted/Betriebsleitung.csv"

#Gets content out of column named Computer (In the first Row)
$ArrComputers = $Temp.Computer


#Name of the File for the Results (Extension will be added automatically)
$NAMEFINISH="Endergebnis"

#Path where to save the Results (the success and failed ones) It's recommmened to save the Files in a sperate Folder
$path=".\FINISH"

#Checks if the Folder (which provided above) exists, if not creating one
If (!(test-path $path))
{
    mkdir $path
}

# Creating whole path combining Filename and PATH
$filename=$path + "\" + $NAMEFINISH + ".csv"

#If File already exists, add Current Day, Month and Time to file name

if (Test-Path $filename) {
    $NAMEFINISH=$NAMEFINISH+(Get-Date -Format "dddd ddMM HHmm")
}

# PATH for failed Computer Results do not edit
$FailedPATHDEF=$path + "\Failed.csv"

# Filename for failed Results (extension will be added later)
$FailedName="FAILED"

#Checks if File already exists, if so then changing the filename to Current Date, Month, and Time
if (Test-Path $FailedPATHDEF) {
$FailedName=$FailedName+(Get-Date -Format "dddd ddMM HHmm")
}


#Interates over every Computer in List
foreach ($Computer in $ArrComputers) 
{

    # Test Connection, if not aviable, skipping to next one and putting hostname into the Failed Results File
    If (!(Test-Connection -ComputerName $Computer -Count 1 -Quiet)) { 
        Write-Host "$Computer not on network."
        "-------------------------------------------------------"

        $failed=[PsCustomObject]@{
        'ComputerName'=$Computer
        'failedon'=$(Get-Date)
        }

        

        $FailedExport=".\FINISH\$($FailedName).csv"

        $failed | Export-CSV $FailedExport -Append -Force -NoTypeInformation

        Continue # Move to next computer
       }

    #Oututs Computer Name for Better Display
    write-host $Computer

    #Little Indicatior to know progress
    write-host "Getting Information about the System"

    #Getting Information about the System from the Remote Computer
    $computerSystem = get-wmiobject Win32_ComputerSystem -Computer $Computer
    

    write-host "Getting Information about the BIOS"
    $computerBIOS = get-wmiobject Win32_BIOS -Computer $Computer
    

    write-host "Getting Information about the Operarting System"
    $computerOS = get-wmiobject Win32_OperatingSystem -Computer $Computer


    write-host "Getting Information about the Processor"
    $computerCPU = get-wmiobject Win32_Processor -Computer $Computer


    write-host "Getting Information about the Drive"
    $computerHDD = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter drivetype=3


    #Searching for Office
    #Note this can take up a while, since it's looking up all installed Programs on the Remote Computer
    #Please be patient

    write-host "Searching Installed Programs for Microsoft Office"
    write-host "This can take a while, please be patient"
    $office=Get-WmiObject -Class Win32_Product -ComputerName $Computer -filter "name like '%Office%'"
    #$office = Get-WmiObject -Class Win32_Product -ComputerName $Computer | Where-Object {$_.Name -Match "Microsoft Office"

    # Optional for getting Drive Compacity, currentlty not working
    #$HDDCAP="{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
    #$HDDSPA="{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"

    write-host "Done, Getting Results"

    #Calculating RAM
    $RAM="{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"

    #Spilting CPU String into Manufactor, Processor and Clockspeed Only on Intel
    $cpu=$computerCPU.Name
    $Split=$cpu.Split("@")
    $GHZ=$Split[1]

    $Rest=$Split[0].split(")")
    $CPU=$Rest[2]
    $Manufactor= $Rest[0].replace("(R","")



    #Converting Object so it can be exported in csv
    $office=$office.Name -join ', '
    $officefinal=$office.Split(",")
    $office=$officefinal[0]
    #$office=$office -join ', '

    #Getting all Informations into one Object, to export them in a csv sheet
    $AllInfos = [PsCustomObject]@{
     'ComputerName'=$computerSystem.PSComputerName
     'Modell'=$computerSystem.SystemFamily
     'lastLogin'=$computerSystem.UserName
     'Manufacturer'=$computerSystem.Manufacturer
     'CPUM:'=$Manufactor
     'CPUP:'=$CPU
     'CPUC:'=$GHZ
     #'HDD Capacity'=$HDDCAP
     #'HDD Space'=$HDDSPA
     'RAM'=$RAM
     'Operating System'=$computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
     'Office'=$office
    }

    $ExportPATH=".\FINISH\$($NAMEFINISH).csv"

    #Exporting the Object into the specified csv file
    $AllInfos | Export-CSV $ExportPATH -Append -Force -NoTypeInformation
    
        #Output the Informations in the console to.
        write-host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
        "-------------------------------------------------------"
        "Manufacturer: " + $computerSystem.Manufacturer
        "Model: " + $computerSystem.Model
        "Serial Number: " + $computerBIOS.SerialNumber
        "CPU: " + $computerCPU.Name
        #"HDD Capacity: "  + "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
        #"HDD Space: " + "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
        "RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
        "Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
        "User logged In: " + $computerSystem.UserName
        "Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        "Office installed:" + $office.Name
        "-------------------------------------------------------"
}

#Pausing the script after finish, added for viewing Results when run from Explorer
write-host "Script finished"
pause