$ResultsHash = [ordered]@{}

$SystemInfo = systeminfo

# Convert systeminfo to a hashtable

$SystemHash = @{}
$SystemInfo -split "`n" | ForEach-Object {
    $Parts = $_ -split ":", 2
    if ($Parts.Count -eq 2) {
        $Key = $Parts[0].Trim()
        $Value = $Parts[1].Trim()
        $SystemHash[$Key] = $Value
    }
}

$ResultsHash["Host Name"] = $SystemHash["Host Name"]
$ResultsHash["OS Name"] = $SystemHash["OS Name"]
# $ResultsHash["OS Version"] = $SystemHash["OS Version"]
# $ResultsHash["OS Manufacturer"] = $SystemHash["OS Manufacturer"]
# $ResultsHash["OS Configuration"] = $SystemHash["OS Configuration"]
# $ResultsHash["OS Build Type"] = $SystemHash["OS Build Type"]
# $ResultsHash["Registered Owner"] = $SystemHash["Registered Owner"]
# $ResultsHash["Registered Organization"] = $SystemHash["Registered Organization"]
# $ResultsHash["Product ID"] = $SystemHash["Product ID"]
# $ResultsHash["Original Install Date"] = $SystemHash["Original Install Date"]
# $ResultsHash["System Boot Time"] = $SystemHash["System Boot Time"]
$ResultsHash["System Manufacturer"] = $SystemHash["System Manufacturer"]
$ResultsHash["System Model"] = $SystemHash["System Model"]
# $ResultsHash["System Type"] = $SystemHash["System Type"]
# $ResultsHash["Processor(s)"] = $SystemHash["Processor(s)"]
$ResultsHash["BIOS Version"] = $SystemHash["BIOS Version"]
# $ResultsHash["Windows Directory"] = $SystemHash["Windows Directory"]
# $ResultsHash["System Directory"] = $SystemHash["System Directory"]
# $ResultsHash["Boot Device"] = $SystemHash["Boot Device"]
# $ResultsHash["System Locale"] = $SystemHash["System Locale"]
# $ResultsHash["Input Locale"] = $SystemHash["Input Locale"]
# $ResultsHash["Time Zone"] = $SystemHash["Time Zone"]
$ResultsHash["Total Physical Memory"] = $SystemHash["Total Physical Memory"]
# $ResultsHash["Available Physical Memory"] = $SystemHash["Available Physical Memory"]
# $ResultsHash["Virtual Memory: Max Size"] = $SystemHash["Virtual Memory: Max Size"]
# $ResultsHash["Virtual Memory: Available"] = $SystemHash["Virtual Memory: Available"]
# $ResultsHash["Virtual Memory: In Use"] = $SystemHash["Virtual Memory: In Use"]
# $ResultsHash["Page File Location(s)"] = $SystemHash["Page File Location(s)"]
# $ResultsHash["Domain"] = $SystemHash["Domain"]
# $ResultsHash["Logon Server"] = $SystemHash["Logon Server"]
# $ResultsHash["Hotfix(s)"] = $SystemHash["Hotfix(s)"]
# $ResultsHash["Network Card(s)"] = $SystemHash["Network Card(s)"]
# $ResultsHash["Virtualization-based security"] = $SystemHash["Virtualization-based security"]
# $ResultsHash["Hyper-V Requirements"] = $SystemHash["Hyper-V Requirements"]

# Retrieve the Windows Product Key
$ProductKey = (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey

$ResultsHash["Windows Product Key"] = $ProductKey

# Get RAM slot information
$memorySlots = Get-CimInstance Win32_MemoryDeviceLocation
$motherBoard = Get-CimInstance win32_baseboard

$RAMSlots = @();

switch ($motherBoard.Product) {
    #find the motherboard models for the most common models and populate manually w/ count of ram slots 
    "0TM99H" { $Totalslots = 2 }
    Default { $Totalslots = 4 }
}

$RAMData = Get-WmiObject Win32_PhysicalMemory |
Select-Object PSComputerName, DeviceLocator, Manufacturer, PartNumber, 
@{ label = "Size/GB"; expression = { $_.Capacity / 1GB } },
Speed, Datawidth, TotalWidth, @{ label = "FreeSlots"; exp = { $Totalslots - $memorySlots.Count } }

foreach ($RAMSlot in $RAMData) {
    if ($RAMSlot.DeviceLocator -like "*DIMM*") {
        $RAMSlots += [PSCustomObject]@{
            DeviceLocator = $RAMSlot.DeviceLocator
            Manufacturer  = $RAMSlot.Manufacturer
            Size_GB       = $RAMSlot."Size/GB"
            Speed_MHz     = $RAMSlot.Speed
        }
    }
}

$ResultsHash["RAM"] = $SystemHash["Total Physical Memory"]

if ($RAMSlots.Count -gt 0) {
    $ResultsHash["RAM Slots"] = $RAMSlots
}
else {
    $ResultsHash["RAM Slots"] = "RAM NOT UPGRADABLE"
}

$CPUDetails = Get-WmiObject -Class Win32_Processor -ComputerName. | Select-Object -Property [a-z]*

$ResultsHash["CPU Name"] = $CPUDetails.Name
$ResultsHash["CPU Cores"] = $CPUDetails.NumberOfCores
$ResultsHash["CPU Threads"] = $CPUDetails.NumberOfLogicalProcessors
$ResultsHash["CPU Max Clock Speed MHz"] = $CPUDetails.MaxClockSpeed

# Battery Health
powercfg /batteryreport /XML /OUTPUT "batteryreport.xml" | Out-Null
Start-Sleep 1
[xml]$BatteryReport = Get-Content -Path "batteryreport.xml"

$BatteryReport.BatteryReport.Batteries |
ForEach-Object {
    $ResultsHash["Battery"] += [PSCustomObject]@{
        BatteryHealth      = "$([math]::floor([int64]$_.Battery.FullChargeCapacity / [int64]$_.Battery.DesignCapacity * 100))%"
        DesignCapacity     = $_.Battery.DesignCapacity
        FullChargeCapacity = $_.Battery.FullChargeCapacity
        CycleCount         = $_.Battery.CycleCount
        Id                 = $_.Battery.id
    }
}

# Storage Details
$PhysicalDisks = Get-PhysicalDisk | Sort-Object Size | Select-Object FriendlyName, Size, MediaType, SpindleSpeed, HealthStatus, OperationalStatus

$ResultsHash["Physical Disks"] = $PhysicalDisks

# Output the results as a well-formatted text file
$OutputPath = "SystemStats.txt"
$Output = @()

$Output += "SYSTEM STATISTICS REPORT"
$Output += "Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$Output += "Made by Euan Bell"
# $Output += "=" * 60

foreach ($Key in $ResultsHash.Keys) {
    $Value = $ResultsHash[$Key]
    
    $Output += "-" * 40
    $Output += "$Key"
    
    if ($Value -is [Array] -and $Value.Count -gt 0 -and $Value[0] -is [PSCustomObject]) {
        # Handle arrays of objects (like RAM Slots, Physical Disks, Battery)
        foreach ($Item in $Value) {
            $Output += ""
            foreach ($Property in $Item.PSObject.Properties) {
                $Output += "  $($Property.Name): $($Property.Value)"
            }
        }
    }
    elseif ($Value -is [String] -or $Value -is [Int] -or $Value -eq $null) {
        # Handle simple values
        $Output += "$Value"
    }
    else {
        # Handle other types
        $Output += "$($Value | Out-String)".Trim()
    }
}

# Write to file and display success message
$Output | Out-File -FilePath $OutputPath -Encoding UTF8
Write-Host "System statistics saved to: $OutputPath" -ForegroundColor Green
