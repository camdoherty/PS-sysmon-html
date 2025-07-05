# --- Configuration ---
# Define the output file names and path
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$outputFolder = Join-Path -Path $PSScriptRoot -ChildPath "reports\$($env:COMPUTERNAME)-$timestamp"
$reportFile = Join-Path $outputFolder "WinHealthReport_Report.html"
$errorsFile = Join-Path $outputFolder "WinHealthReport_Errors.html"

# Define services to check. Add your own critical services here.
$servicesToCheck = @(
    "spooler",
    "wuauserv",
    "WinRM",
    "NonExistentService" # Included to demonstrate error handling
)

# Define thresholds for warnings
$diskWarningThreshold = 90 # Percentage
$diskCriticalThreshold = 95 # Percentage
$cpuWarningThreshold = 90 # Percentage
$memoryWarningThreshold = 90 # Percentage


# --- HTML & CSS Styling ---
$htmlHeader = @"
<html>
<head>
<style>
body { font-family: Calibri, sans-serif; }
h1, h2 { color: #2E4053; }
table {
  border: 1px solid black;
  border-collapse: collapse;
  width: 80%;
  margin-bottom: 20px;
}
th, td {
  padding: 8px;
  text-align: left;
  border-bottom: 1px solid #ddd;
}
th {
  background-color: #4CAF50;
  color: white;
}
tr:nth-child(even) { background-color: #f2f2ff; }
.section-header { font-size: 1.5em; color: #2E4053; margin-top: 20px;}
.warning { background-color: #FFC300; font-weight: bold; }
.critical { background-color: #C70039; color: white; font-weight: bold; }
a { color: #0000EE; }
</style>
</head>
<body>
<h1>System Health Report</h1>
"@

$htmlFooter = @"
</body>
</html>
"@

# Create the output folder if it does not exist
if (-not (Test-Path -Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
}

# --- Data Collection ---

# Create an array to hold all the HTML fragments
$reportFragments = @()

# --- 1. General System Information ---
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$computerInfo = Get-CimInstance -ClassName Win32_ComputerSystem
$uptime = (Get-Date) - $osInfo.LastBootUpTime
$generalInfo = [PSCustomObject]@{
    "Computer Name"      = $env:COMPUTERNAME
    "Report Generated"   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "System Uptime"      = "{0:N0} days, {1:N0} hours, {2:N0} minutes" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
    "Operating System"   = $osInfo.Caption
    "OS Architecture"    = $osInfo.OSArchitecture
}
$reportFragments += "<h2>System Information</h2>"
$reportFragments += $generalInfo | ConvertTo-Html -As 'List' -Fragment

# --- 2. Event Viewer Errors (with clickable link) ---
$reportFragments += "<h2>Key Metrics</h2>"
$eventErrors = Get-WinEvent -FilterHashtable @{LogName = "System"; Level = 1, 2; StartTime = (Get-Date).AddHours(-24) } -ErrorAction SilentlyContinue

$eventErrorCount = if ($null -ne $eventErrors) { $eventErrors.Count } else { 0 }
if ($eventErrorCount -gt 0) {
    # Create a link to the detailed errors file
    $errorLink = "<a href='$(Split-Path -Path $errorsFile -Leaf)'>$eventErrorCount (Click for details)</a>"

    # Generate the separate errors page
    $errorDetails = $eventErrors | Select-Object TimeCreated, Id, LevelDisplayName, Message | ConvertTo-Html -Head $htmlHeader -Body "<h1>System Errors (Last 24 Hours) on $($env:COMPUTERNAME)</h1>"
    $errorDetails | Out-File -FilePath $errorsFile
}
else {
    $errorLink = "0"
}

# Use an ordered dictionary to control the order of rows in the summary table
$summaryData = [ordered]@{}
$summaryData["Event Viewer Errors (last 24 hours)"] = $errorLink


# --- 3. Memory Utilization ---
$memory = $osInfo
$totalMemoryGB = [math]::Round($memory.TotalVisibleMemorySize / 1MB, 2)
$freeMemoryGB = [math]::Round($memory.FreePhysicalMemory / 1MB, 2)
$usedMemoryGB = $totalMemoryGB - $freeMemoryGB
$memoryPercentUsed = [math]::Round(($usedMemoryGB / $totalMemoryGB) * 100, 2)
$memoryStatus = if ($memoryPercentUsed -ge $memoryWarningThreshold) { "class='critical'" } else { "" }
$summaryData["Memory Utilization"] = "<td $memoryStatus>$memoryPercentUsed% Used ($usedMemoryGB GB of $totalMemoryGB GB)</td>"


# --- 4. CPU Utilization ---
# Get a single snapshot of CPU load. Note: This is instantaneous.
$cpuLoad = (Get-CimInstance -ClassName Win32_Processor).LoadPercentage
$cpuStatus = if ($cpuLoad -ge $cpuWarningThreshold) { "class='critical'" } else { "" }
$summaryData["CPU Utilization"] = "<td $cpuStatus>$cpuLoad%</td>"


# Create HTML table from the summary data
$summaryTable = "<table>"
foreach ($key in $summaryData.Keys) {
    if ($key -eq "Memory Utilization" -or $key -eq "CPU Utilization") {
        # These have custom <td> tags for coloring, so don't wrap them in another <td>
        $summaryTable += "<tr><td>$key</td>$($summaryData[$key])</tr>"
    } else {
        $summaryTable += "<tr><td>$key</td><td>$($summaryData[$key])</td></tr>"
    }
}
$summaryTable += "</table>"
$reportFragments += $summaryTable


# --- 5. Disk Usage ---
$reportFragments += "<h2>Disk Usage</h2>"
$diskUsage = Get-CimInstance -Class Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    $sizeGB = [math]::Round($_.Size / 1GB, 2)
    $freeGB = [math]::Round($_.FreeSpace / 1GB, 2)
    $percentUsed = if ($_.Size -gt 0) { [math]::Round(($_.Size - $_.FreeSpace) / $_.Size * 100, 2) } else { 0 }
    
    $status = "OK"
    if ($percentUsed -ge $diskCriticalThreshold) {
        $status = "CRITICAL"
    } elseif ($percentUsed -ge $diskWarningThreshold) {
        $status = "Warning"
    }

    [PSCustomObject]@{
        'Drive'        = $_.DeviceID
        'Label'        = $_.VolumeName
        'Total Size (GB)' = $sizeGB
        'Free Space (GB)' = $freeGB
        'Used (%)'     = $percentUsed
        'Status'       = $status
    }
}
$reportFragments += $diskUsage | ConvertTo-Html -Fragment


# --- 6. Temperature Information (Fixed and more robust) ---
$reportFragments += "<h2>Hardware Temperature</h2>"
try {
    # MSAcpi_ThermalZoneTemperature is more common than CIM_TemperatureSensor.
    # The temperature is in tenths of a Kelvin.
    $temperatureInfo = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop
    
    if ($null -ne $temperatureInfo) {
        $tempReadings = $temperatureInfo | ForEach-Object {
            [PSCustomObject]@{
                'Sensor'        = $_.InstanceName
                'Temperature (C)' = [math]::Round(($_.CurrentTemperature / 10) - 273.15, 2)
            }
        }
        $reportFragments += $tempReadings | ConvertTo-Html -Fragment
    } else {
        $reportFragments += "<p>No temperature sensors were found on this system.</p>"
    }
}
catch {
    # This catch block runs if the WMI class doesn't exist.
    $reportFragments += "<p>Temperature data is not available on this system (WMI class not found).</p>"
}


# --- 7. Top Resource-Consuming Processes ---
$reportFragments += "<h2>Top 5 Processes by CPU</h2>"
$topCpu = Get-Process | Sort-Object -Property CPU -Descending | Select-Object -First 5 | Select-Object Name, Id, @{N = 'CPU Time (s)'; E = { [math]::Round($_.CPU, 2) } }, WorkingSet
$reportFragments += $topCpu | ConvertTo-Html -Fragment

$reportFragments += "<h2>Top 5 Processes by Memory</h2>"
$topMem = Get-Process | Sort-Object -Property WorkingSet -Descending | Select-Object -First 5 | Select-Object Name, Id, @{N = 'Memory (MB)'; E = { [math]::Round($_.WorkingSet / 1MB, 2) } }
$reportFragments += $topMem | ConvertTo-Html -Fragment


# --- 8. Monitored Service Status ---
$reportFragments += "<h2>Monitored Service Status</h2>"
$serviceStatus = foreach ($serviceName in $servicesToCheck) {
    try {
        $service = Get-Service -Name $serviceName -ErrorAction Stop
        [PSCustomObject]@{
            'Service Name' = $service.DisplayName
            'Status'       = $service.Status
        }
    }
    catch {
        [PSCustomObject]@{
            'Service Name' = $serviceName
            'Status'       = 'NOT FOUND'
        }
    }
}
$reportFragments += $serviceStatus | ConvertTo-Html -Fragment


# --- Assemble and Save the Report ---
$htmlBody = $reportFragments -join "`n"

# Replace default TD and TR with class-based ones for coloring
$htmlBody = $htmlBody -replace '<td>CRITICAL</td>', '<td class="critical">CRITICAL</td>'
$htmlBody = $htmlBody -replace '<td>Warning</td>', '<td class="warning">Warning</td>'
$htmlBody = $htmlBody -replace '<td>Stopped</td>', '<td class="critical">Stopped</td>'
$htmlBody = $htmlBody -replace '<td>NOT FOUND</td>', '<td class="warning">NOT FOUND</td>'

$htmlContent = $htmlHeader + $htmlBody + $htmlFooter

$htmlContent | Out-File -FilePath $reportFile

Write-Host "System report generated successfully at '$reportFile'"
if ($eventErrorCount -gt 0) {
    Write-Host "Error details report generated at '$errorsFile'"
}
