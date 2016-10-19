##########################################################################################
# Name: Exchange-ReceiveConnectorLogAnalyser.ps1
# Author: Adrian Begg (adrian.begg@ehloworld.com.au)
#
# Date: 1/10/2016
#
# Purpose: The purpose of this script is to generate a summary report from the SMTP 
# Transport logs of hosts that have connected to Receive Connectors. This is intended to be 
# used when migrating/decommission of an Exchange server or connector to determine what
# (e.g.. Rouge MFD's or Applications where someone just put in an IP address 20 years ago) 
# is still connecting. Outputs a simple CSV file.
##########################################################################################
# REQUIREMENTS:
# 1) Protocol Logging must be enabled on the connectors you wish to assess - Refer to https://technet.microsoft.com/en-us/library/bb124531(v=exchg.150).aspx for more 
# information on enabling Protocol Logging and setting the path.
# 
# 2) You must have permission to the NTFS folder hosting Protocol Logs 
#
# ASSUMPTIONS:
# 1) It is assumed that this script is being executed from the local Exchange Server

Function Save-FileName([string] $initialDirectory){   
	$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
	$SaveFileDialog.initialDirectory = $initialDirectory
	$SaveFileDialog.filter = “CSV (*.csv)|*.csv|All files (*.*)|*.*”
	$SaveFileDialog.ShowDialog() | Out-Null
	return $SaveFileDialog.Filename
}

# MAIN
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
# The path to the transport logs to process
[string] $strExchangeLogPath = "$env:ExchangeInstallPath\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive"
# A temp file to store a concatenated  version of all of the available log files
[string] $strMasterLog = "$env:temp\consolidatedSmtpReceive.log"

# The output file to dump the report
[System.Windows.Forms.MessageBox]::Show("Please select the filename/path of the output CSV file.") | Out-Null
[string] $strOutputCSV = Save-FileName $PSScriptRoot

# Check if a file was specified
if([string]::IsNullOrEmpty($strOutputCSV)){
	throw "No Output CSV has been specified. Please specify an output filename and try again."
}

# Next check if the Log Path exists
if(!(Test-Path $strExchangeLogPath)){
	throw "The Protocol Log Path can not be found. Please verify that Exchange is installed and retry"
}

# Create the header in the temp file for processing
Out-File -FilePath $strMasterLog -InputObject "date-time,connector-id,session-id,sequence-number,local-endpoint,remote-endpoint,event,data,context" -Encoding ascii

$colLogFiles = Get-ChildItem $strExchangeLogPath -Filter *.log
foreach($objLogFile in $colLogFiles){
	$logContent = Get-Content $objLogFile.FullName
	foreach($objLogLine in $logContent){
		if(!$objLogLine.StartsWith("#")){
			Add-Content -Path $strMasterLog $objLogLine 
		}
	}
}

# Now process the unique objects
$objLogCSV = Import-CSV $strMasterLog

# Remove the Port from the Source IP address Property
foreach($objEntity in $objLogCSV){
	$objEntity."remote-endpoint" = $objEntity."remote-endpoint".Substring(0,$objEntity."remote-endpoint".IndexOf(":"))
}
$colSummary = $objLogCSV | Group-Object -Property "remote-endpoint","connector-id"

#ArrayList to store the report results for Export to CSV
$colResults = New-Object -TypeName System.Collections.ArrayList
foreach($objHost in $colSummary){
	# Setup the object to store the values for each host
	[string] $IPAddress = $objHost.Name.Substring(0,$objHost.Name.IndexOf(","))
	[string] $ConnectorId = $objHost.Name.Substring(($objHost.Name.IndexOf(",")+2),(($objHost.Name.Length) - ($objHost.Name.IndexOf(",")+2)))
	[string] $DNSHostName = ""
	
	# Next attempt a reverse DNS lookup for the IP
	try{
		$DNSHostName = [System.Net.Dns]::GetHostEntry($IPAddress).HostName
	} catch {
		 # Do nothing just need to catch the exception and continue
	}
	$objHostResult = New-Object System.Management.Automation.PSObject
	$objHostResult | Add-Member Note* ConnectorId $ConnectorId
	$objHostResult | Add-Member Note* IPAddress $IPAddress
	$objHostResult | Add-Member Note* HostName $DNSHostName
	$objHostResult | Add-Member Note* Count $objHost.Count
	$colResults.Add($objHostResult) > $null
}
# Finally export the report to a CSV and clean up the input object
$colResults | Export-CSV -Path $strOutputCSV -NoTypeInformation
Remove-Item $strMasterLog -Force

[System.Windows.Forms.MessageBox]::Show("Processing Complete :)") | Out-Null
# Quit
Exit