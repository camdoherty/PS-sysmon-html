# Define the output file name and path
$outputFile = "C:\Temp\SystemReport.html"

# Define the HTML header and style
$htmlHeader = @"
<html>
<head>
<style>
table, th, td {
  border: 1px solid black;
  border-collapse: collapse;
}
th, td {
  padding: 5px;
  text-align: left;
}
</style>
</head>
<body>
<h1>System Report</h1>
"@

# Define the HTML footer
$htmlFooter = @"
</body>
</html>
"@

# Create an empty array to store the table rows
$tableRows = @()

# Get the event viewer errors from the last 24 hours
$eventErrors = Get-WinEvent -FilterHashtable @{LogName="System"; Level=1,2; StartTime=(Get-Date).AddHours(-24)}

# Add a table row with the event viewer errors count
$tableRows += "<tr><td>Event Viewer Errors (last 24 hours)</td><td>$($eventErrors.Count)</td></tr>"

# Get the disk(s) usage information
$diskUsage = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"

# Loop through each disk and add a table row with the disk usage percentage
foreach ($disk in $diskUsage) {
    $diskPercent = [math]::Round(($disk.Size - $disk.FreeSpace) / $disk.Size * 100, 2)
    $tableRows += "<tr><td>Disk Usage ($($disk.DeviceID))</td><td>$diskPercent%</td></tr>"
}

# Get the CPUs usage information
$cpuUsage = Get-WmiObject -Class Win32_Processor

# Loop through each CPU and add a table row with the CPU usage percentage
foreach ($cpu in $cpuUsage) {
    $cpuPercent = [math]::Round($cpu.LoadPercentage, 2)
    $tableRows += "<tr><td>CPU Usage ($($cpu.DeviceID))</td><td>$cpuPercent%</td></tr>"
}

# Get the memory utilization percentage
$memoryUsage = Get-WmiObject -Class Win32_OperatingSystem

# Calculate the memory utilization percentage
$memoryPercent = [math]::Round(($memoryUsage.TotalVisibleMemorySize - $memoryUsage.FreePhysicalMemory) / $memoryUsage.TotalVisibleMemorySize * 100, 2)

# Add a table row with the memory utilization percentage
$tableRows += "<tr><td>Memory Utilization</td><td>$memoryPercent%</td></tr>"

# Get the CPU and GPU temperature information
$temperatureInfo = Get-WmiObject -Class MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"

# Loop through each temperature sensor and add a table row with the temperature value in Celsius
foreach ($temp in $temperatureInfo) {
    # Convert the temperature value from Kelvin to Celsius
    $tempCelsius = [math]::Round(($temp.CurrentTemperature / 10) - 273.15, 2)
    # Check if the sensor is for CPU or GPU
    if ($temp.InstanceName -like "*CPU*") {
        $tableRows += "<tr><td>CPU Temperature</td><td>$tempCelsius°C</td></tr>"
    }
    elseif ($temp.InstanceName -like "*GPU*") {
        $tableRows += "<tr><td>GPU Temperature</td><td>$tempCelsius°C</td></tr>"
    }
}

# Join the table rows with a new line character
$tableBody = $tableRows -join "`n"

# Create the HTML table with the table body
$htmlTable = "<table>$tableBody</table>"

# Create the HTML content by concatenating the header, table and footer
$htmlContent = $htmlHeader + $htmlTable + $htmlFooter

# Write the HTML content to the output file
$htmlContent | Out-File -FilePath $outputFile

```.
