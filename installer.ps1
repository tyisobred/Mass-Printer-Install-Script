
<#
    Creator: Tyler Melcavage
    Version: 1.6 04/23/2025
    
    Purpose: Add printer to computers listed  in C:\temp\PRINTER_computers.txt

    At this time, only Ricoh and HP printers that are able to use the universal drivers are supported.
    V1.1 Make sure that printer is online/correct name before installing
    V1.2 Added additional options
    V1.3 Added progress bar status
    V1.4 Added ability to add driver to driver store using psexec if not present.
    v1.5 Added check for existing Printer name and options to pick which PCL6 driver version to install.
    v1.6 Removed remote installation of drivers. Provided options to install known installed drivers or most updated of a model

#>

 # No Debug
 $DEBUGPREFVAL = "SilentlyContinue"
 # Yes Debug
 #$DEBUGPREFVAL = "Continue"


 #$DebugPreference = "Continue"
 $DebugPreference = $DEBUGPREFVAL
 
 
function Get-UserSelectedDriver {
    $DebugPreference = $DEBUGPREFVAL

        $printerChoices = @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&Ricoh", "Ricoh Universal Driver PCL6")
            [System.Management.Automation.Host.ChoiceDescription]::new("&HP", "HP Universal Printing PCL6")
        )
    
        $ricohDriverChoices = @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&1 - Most Updated", "MostUpdatedRICOHDriver")
            [System.Management.Automation.Host.ChoiceDescription]::new("&2 - v4.40", "Ricoh Universal Driver V4.40.0.0")
            [System.Management.Automation.Host.ChoiceDescription]::new("&3 - v4.37", "Ricoh Universal Driver V4.37.0.0")
        )

        $hpDriverChoices = @(
            [System.Management.Automation.Host.ChoiceDescription]::new("&1 - Most Updated", "MostUpdatedHPDriver")
            [System.Management.Automation.Host.ChoiceDescription]::new("&2 - v7.1.0", "HP Universal Driver V7.1.0.25570")
        )

        $returnedDriver = @{}

        Write-Host "If the driver selected below is not already installed on the device, it will be skipped"

        $printerChoices = $Host.UI.PromptForChoice("Which model of printer is being installed? ","",$printerchoices,-1)
     
        # User choose Ricoh
        if($printerChoices -eq 0) {
            $printerModelName = "Ricoh"
            $driverChoices = $Host.UI.PromptForChoice("Which Ricoh Driver would you like to install? ","",$ricohDriverChoices,0)

            switch ($driverChoices) {
                0 {
                    $returnedDriver.Name = $null
                    $returnedDriver.Model = "RICOH PCL6 UniversalDriver"
                    $returnedDriver.Static = $false
                }
                1 {
                    $returnedDriver.Name = 'RICOH PCL6 UniversalDriver V4.40'
                    $returnedDriver.Model = "RICOH PCL6 UniversalDriver"
                    $returnedDriver.Static = $true
                }
                2 {
                    $returnedDriver.Name = 'RICOH PCL6 UniversalDriver V4.37'
                    $returnedDriver.Model = "RICOH PCL6 UniversalDriver"
                    $returnedDriver.Static = $true
                }
            }
        }

        # User choose HP
        if($printerChoices -eq 1) {
            $printerModelName = "HP"
            $driverChoices = $Host.UI.PromptForChoice("Which HP Driver would you like to use? ","",$hpDriverChoices,0)

            switch ($driverChoices) {
                0 {
                    $DriverName = $null
                    $returnedDriver.Model = "HP Universal Printing PCL 6"
                    $returnedDriver.Static = $false
                }
                1 {
                    $DriverName = 'HP Universal Printing PCL 6 (v7.1.0)'
                    $returnedDriver.Model = "HP Universal Printing PCL 6"
                    $returnedDriver.Static = $true
                }
            }
        }
        
        
        Write-Debug "Driver Choosen: $($returnedDriver.Name)"

    return (, $returnedDriver)
}

function Get-MostUpdatedDriver {
    
    param (
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [Parameter(Mandatory)]
        [string]$DriverManufacture
    )
    $DebugPreference = $DEBUGPREFVAL
    
    If($DriverManufacture -ieq "HP Universal Printing PCL 6") {
        Write-Debug "Printer is HP - Append Universal"
        $DriverManufacture = "HP Universal Printing PCL 6"
    }
    ElseIf ($DriverManufacture -ieq "RICOH PCL6 UniversalDriver") {
        Write-Debug "Printer is Ricoh - Append Universal"
        $DriverManufacture = "RICOH PCL6 UniversalDriver"
    } 
    Else {
        Write-Debug "DriverManufacture: $DriverManufacture"
        Write-Error "Parameter for Get-MostUpdatedDriver is incorrect: $DriverManufacture"
        pause
        exit
    }

    $driversInstalled = (Get-PrinterDriver -ComputerName $ComputerName -Name "*$DriverManufacture*" | Select-Object Name, DriverVersion, Manufacturer)

    try{
        $updatedDriverItem = $driversInstalled[0]
    }
    catch {
        Write-Debug "$ComputerName does not have a driver matching: $DriverManufacture"
        Write-Debug $_               
        return (, $null)
    }

    ForEach ($driver in $driversInstalled) {
        Write-Debug "Current Object $($driver.Name) Version: $($driver.DriverVersion)"
        Write-Debug "Compared Object $($updatedDriverItem.Name) Version: $($updatedDriverItem.DriverVersion)"
        
        Write-Debug "Current Highest Version: $($updatedDriverItem.DriverVersion)"
        Write-Debug "Compared Version: $($driver.DriverVersion)"

        try {
        if(([uint64]($driver.DriverVersion)) -gt ([uint64]($updatedDriverItem.DriverVersion))) {
            $updatedDriverItem = $driver
            
            Write-Debug "Most up to date driver installed set to: $updatedDriverItem"

        }
        } catch {
            Write-Debug "Error comparing drivers: $updatedDriverItem"
            Write-Debug $_
        }
    }

    Write-Debug "Returning highest verion driver for $ComputerName : $updatedDriverItem"
    return (, $updatedDriverItem)
  
}

function Get-HostStatus {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName
    )
    $DebugPreference = $DEBUGPREFVAL

    $pingResult = Test-Connection -BufferSize 32 -Count 1 -ComputerName $ComputerName -Quiet
    return $pingResult
}

function Add-RemotePrinterPort {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [Parameter(Mandatory)]
        [string]$PortName
    )
    $DebugPreference = $DEBUGPREFVAL

    Write-Debug "Testing if port already exist"
    $printerPortExists = Get-PrinterPort -Name $PortName -ComputerName $ComputerName -ErrorAction SilentlyContinue

    $portStatus = "Unknown"
    if (-not $printerPortExists) {
        Write-Debug "Port is not present. Adding Port to remote computer" 
        Write-Information "Adding TCP/IP Port on $ComputerName."
        Add-PrinterPort -Name $PortName -PrinterHostAddress $PortName -ComputerName $ComputerName
        $portStatus = "Port Added"
    } 
    Else {
        Write-Debug "Port already exist. Using existing Port"
        Write-Information "Port already exist on $ComputerName. Using existing port"
        $portStatus = "Port Already Exist"
    }

    return $portStatus
}

function Add-RemotePrinterDevice {
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,
        [Parameter(Mandatory)]
        [string]$PrinterName,
        [Parameter(Mandatory)]
        [string]$PortName,
        [Parameter(Mandatory)]
        [string]$DriverName
    )
    $DebugPreference = $DEBUGPREFVAL

    Write-Debug "Adding Printer to remote computer"
    $printerNameExists = Get-Printer -Name $PrinterName -ComputerName $Computer -ErrorAction SilentlyContinue
    
    Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "80% Complete:" -PercentComplete 80 -CurrentOperation "Adding Printer"

    $printerInstallStatus = "Unknown"
    if (-not $printerNameExists) {
        Write-Debug "Printer name is not present. Adding Printer to remote computer"
        try {
            Add-Printer -name $PrinterName -PortName $PortName -DriverName "$DriverName" -ComputerName $ComputerName
        
            $printerInstallStatus = "Installed"
        } catch {
            Write-Debug "Error Adding Printer ($PrinterName) on $ComputerName"
            Write-Debug $_
            $printerInstallStatus = "Failed"
        }
    } 
    Else {
        Write-Debug "Printer Name already exist. Skipping computer"
        Write-Information "Printer name already exist on $ComputerName. Skipping $ComputerName"
        $printerInstallStatus = "Skipped - Printer Name already exist"
    }

    return $printerInstallStatus
}


$Results=@()

$pcFilePath = "C:\temp\PRINTER_computers.txt"
try {
    while ( $true ) {

        $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")


        $statusDataPath = "C:\temp\printerStatus_$(get-date -f yyyy-MM-dd_HHmmss).csv"
        $Computers = Get-Content -Path $pcFilePath

        Write-Debug "Reading Prompt: Port Name"
        $PortName = Read-Host -Prompt 'Input Port Name'

        Write-Debug "Testing Port Connectivitity"
        $PrinterStatus = Get-HostStatus -ComputerName $PortName
        $tempPortName = $PortName

        While($PrinterStatus -ne 'True' ) {
            #Write-Error -Message "Printer Unreachable: Make sure printer is online" -Category ConnectionError
            Write-Warning "To continue with offline port, please re-enter the port name exactly"
            Write-Debug "Port is offline. Reading Prompt: Port Name"
            $PortName = Read-Host -Prompt 'Input Port Name'

            if($PortName -eq $tempPortName)
            {
                Write-Debug "User re-entered offline port. Bypassing Connectivity Check and Continuing"
                $PrinterStatus = "True"
            }
            else {
                Write-Debug "Testing updated Port Connectivitity"
                $PrinterStatus = Get-HostStatus -ComputerName $PortName
                $tempPortName = $PortName
            }
        }



        Write-Debug "Port is Online"
        Write-Debug "Read Prompt: Printer Name"
        if(!($PrinterName = Read-Host -Prompt "Input Printer Name [$PortName]")) { $PrinterName = $PortName }
        
        Write-Host "`n"

        $numPC = $Computers.Length
        $currentPC = 0
        $percentComplete = 0
        $numFailed = 0
        $itemResultID = 0

        $driverUserChoice = Get-UserSelectedDriver

        ForEach ($Computer in $Computers) {
            Write-Progress -Activity "Adding Printers to multiple computers" -Id 0 -Status "$percentComplete% Complete ($numFailed Offline):" -PercentComplete $percentComplete
            Write-Debug "Testing Computer Connectivitity"

            $Status = Get-HostStatus -ComputerName $Computer

            Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "0% Complete:" -PercentComplete 0 -CurrentOperation "Adding Driver"
            If($Status -eq 'True') {
                Write-Debug "Computer Online"
    

                #Add Driver
                Write-Debug "Adding Driver to remote computer"
                Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "16% Complete:" -PercentComplete 16 -CurrentOperation "Adding Driver"

                if ($driverUserChoice.Static -eq $true) {
                    #Same driver for each computer - check if driver is installed aldreay
                    if(-not (Get-PrinterDriver -ComputerName $Computer | Where-Object { $_.Name -eq "$($driverUserChoice.Name)" }))
                    {
                        $DriverName = $null
                    }
                    else {
                        $DriverName = $driverUserChoice.Name
                    }
                }
                else {
                    #Driver name changes to most updated on each machine

                    $DriverSelectedItem = Get-MostUpdatedDriver -ComputerName $Computer -DriverManufacture $driverUserChoice.Model
                    $DriverName = $DriverSelectedItem.Name
                }


                    

                if(-not($DriverName)) {
                    Write-Debug "$Computer is missing driver. Skipping computer."
                    $numFailed++
                    $PrinterInfo = @{
                        ComputerName = $Computer
                        Status = "Skipped - Missing Drivers"
                        PrinterName = $PrinterName
                        PrinterPort = $PortName
                        PrinterDriver = "Missing Driver"
                        PortStatus = $portStatus
                    }

                    $Results += New-Object psobject -Property $PrinterInfo
                    Write-Progress -Activity "Adding Printers to $Computer" -Id 1 -ParentId 0 -Status "100% Complete:" -PercentComplete 100 -CurrentOperation "Skipped"
                    Continue
                } else {
                    Write-Debug "Driver already in driver store. Using Existing Drivers"
                    Write-Debug "DriverName Variable set to $DriverName"
                }

                Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "30% Complete:" -PercentComplete 30 -CurrentOperation "Adding Driver"
                Add-PrinterDriver -Name "$DriverName" -ComputerName $Computer




                #Add Port
                Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "33% Complete:" -PercentComplete 33 -CurrentOperation "Adding Port"
                $portStatus = Add-RemotePrinterPort -ComputerName $Computer -PortName $PortName

                # Add Printer
                Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "66% Complete:" -PercentComplete 66 -CurrentOperation "Installing Printer"
                $printerInstallationStatus = Add-RemotePrinterDevice -ComputerName $Computer -PrinterName $PrinterName -PortName $PortName -DriverName $DriverName
            
                Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "95% Complete:" -PercentComplete 95 -CurrentOperation "Saving Status"
                $PrinterInfo = @{
                        ComputerName = $Computer
                        Status = $printerInstallationStatus
                        PrinterName = $PrinterName
                        PrinterPort = $PortName
                        PrinterDriver = $DriverName
                        PortStatus = $portStatus
                    }

            
                $Results += New-Object psobject -Property $PrinterInfo 
                Write-Progress -Activity "Adding Printer to $Computer" -Id 1 -ParentId 0 -Status "100% Complete:" -PercentComplete 100 -CurrentOperation "Complete"
            } Else { 
                Write-Debug "$Computer is OFFLINE. Skipping computer."
                $numFailed++
                $PrinterInfo = @{
                    ComputerName = $Computer
                    Status = "PC OFFLINE"
                    PrinterName = $PrinterName
                    PrinterPort = $PortName
                    PrinterDriver = "PC OFFLINE"
                    PortStatus = "PC OFFLINE"
                }

                $Results += New-Object psobject -Property $PrinterInfo
                Write-Progress -Activity "Adding Printers to $Computer" -Id 1 -ParentId 0 -Status "100% Complete:" -PercentComplete 100 -CurrentOperation "Skipped"
            }

            $currentPC = $currentPC + 1
            $percentComplete = [int](($currentPC / $numPC) * 100) - 1
            Write-Debug "Moving to next Computer on list."
        }   
        
        
        Write-Debug "List of computers completed."
        
        Write-Progress -Activity "Adding Printers to multiple computers" -Id 0 -Status "$percentComplete% Complete ($numFailed Offline):" -PercentComplete $percentComplete -CurrentOperation "Checking for More Printers"
       

        $NumSuccessful = ($Results | Where-Object { $_.Status -EQ "Installed" }).Count
        $NumSkipped = ($Results | Where-Object { $_.Status -like "*Skipped*" -or $_.Status -like "*Failed*" }).Count
        $NumOffline = ($Results | Where-Object { $_.Status -like "*OFFLINE*" }).Count

        Write-Host "`nTotal Installed:"
        Write-Host $NumSuccessful

        Write-Host "Total Skipped:"
        Write-Host $NumSkipped

        Write-Host "Total Offline:"
        Write-Host $NumOffline
        
        $choice = $Host.UI.PromptForChoice("Repeat the script? ","",$choices,0)
        if ( $choice -ne 0 ) {
            Write-Debug "User choose not to repeat script."
            Write-Progress -Activity "Adding Printers to multiple computers" -Id 0 -Status "100% Complete ($numFailed Offline):" -PercentComplete 100
            break
        }

        Write-Host "`n"

        Write-Progress -Activity "Adding Printers to multiple computers" -Id 0 -Status "100% Complete ($numFailed Offline):" -PercentComplete 100

        Write-Debug "Repeating Script."
    }
}
finally {
    Write-Debug "Creating CSV Report" 
    Write-Progress -Activity "Adding Printers to multiple computers" -Id 2 -Status "20% Complete" -PercentComplete 20 -CurrentOperation "Generating Report"
    $Results | Select-Object ComputerName, Status, PrinterName, PrinterPort, PrinterDriver, PortStatus| Export-Csv -notypeinformation -Path $statusDataPath
    Write-Progress -Activity "Adding Printers to multiple computers" -Id 2 -Status "100% Complete" -PercentComplete 100 -CurrentOperation "Generating Report"
    Write-Host "Report saved at $statusDataPath"
}
