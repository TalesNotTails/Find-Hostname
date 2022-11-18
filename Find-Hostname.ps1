# Importing required module
Import-Module Import-Excel

# Path to file with static IPs
$ipList = "\path\to\file"

# creating object for AD DNS entries
$dnsList = get-DnsServerResourceRecord -ZoneName "colonial.k12.de.us" -ComputerName colbiggskdc1 -RRType A

## find each sheet in the workbook
$sheets = (Get-ExcelSheetInfo -Path $ipList).Name
## read each sheet and create a CSV file with the same name
foreach ($sheet in $sheets) {
	Import-Excel -WorksheetName $sheet -Path $ipList
}