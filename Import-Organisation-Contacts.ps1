##########################################################################################
# Name: Import-Organisation-Contacts.ps1
# Author: Adrian Begg (adrian.begg@ehloworld.com.au)
#
# Date: 15/02/2017
# Purpose: The purpose of the script is to create Mail Contacts in an organisation
# from user objects in a foriegn Active Directory forest
#
# The primary use case for this is two seperate Active Directory forests with Partner Organisations 
# where the Exchange Organisations must be kept seperate (eg. for legal reasons) however users wish
# 
# Requires the Active Directory modules for Windows PowerShell installed on the executing machine
##########################################################################################

# The URI of the Exchange Server PowerShell Service
$ExchangeServer = "https://labex1.pigeonnuggets.com/powershell"
 
# The file path to the CSV with the objects to import
[string] $strImportObjectsCSV = "D:\Temp\users.csv"

# An audit log of updates made to the directory via the script
[string] $strAuditLog = "AuditLog.txt"

# The DN of the OU which the syncronised objects should be placed
[string] $strTargetOUDN = "OU=Contacts,DC=pigeonnuggets,DC=com"
[string] $strCompany = "Contoso"
[string] $strCustomAttribute1 = "ContosoContact" # Can be used for Address Book Policy or left blank

# Input checking routines
if(([string]::IsNullOrEmpty($strImportObjectsCSV)) -or (!(Test-Path $strImportObjectsCSV))){
	throw "No CSV file has been specified containing the objects or the file does not exist. Please specify an input filename and try again."
}

try{
	Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$strTargetOUDN'" > $nul
} catch {
	throw "The Org Unit provided does not currently exist; please create the OU before attempting to run this process"
}

Add-Content $strAuditLog "$(Get-Date) - INFO :: Execution Stated -- Input file: $strImportObjectsCSV  -- Target OU: $strTargetOUDN"

# Setup the Connection to the Exchange Server and load Active Directory Module
try{
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeServer -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $ExchangeSession -AllowClobber
} catch {
    throw "An error occured loading the modules for administering Exchange. Unable to continue"
}
Import-Module activedirectory

# Load the CSV
$objCSVImport = Import-Csv $strImportObjectsCSV

# Step 1. Determine which objects exist in the organisation already however do not exist in the provided list
$colContactObjects = Get-ADObject -SearchBase $strTargetOUDN -LDAPFilter "objectClass=Contact" -Properties givenName,sn,displayName,name,mail,department,telephoneNumber,title,company

foreach($objCurrent in $colContactObjects){
	# Search for the object in the CSV
	$objIntermediate = $objCSVImport | ?{$_.name -eq $objCurrent.name}
	
	# If the Object does not exist in the CSV it should be removed
	if($objIntermediate.Count -eq 0){
		Add-Content $strAuditLog "$(Get-Date) - INFO :: Removed contact that does not exist in Input file -- $($objCurrent.Name)"
		$objCurrent.Name | Remove-MailContact -Confirm:$False
	} else {
		# Check if any of the attributes on the object have changes and update the object if they have
		if(($objCurrent.givenName -ne $objIntermediate.givenName) -and !([string]::IsNullOrEmpty($objCurrent.givenName))){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Updated Given Name for $($objCurrent.Name) OLD VALUE: $($objCurrent.givenName) NEW VALUE: $($objIntermediate.givenName)"
			Get-MailContact $objCurrent.Name | Set-Contact -FirstName $objIntermediate.givenName
		}
		if(($objCurrent.sn -ne $objIntermediate.sn) -and !([string]::IsNullOrEmpty($objCurrent.sn))){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Updated Surname for $($objCurrent.Name) OLD VALUE: $($objCurrent.sn) NEW VALUE: $($objIntermediate.sn)"
			Get-MailContact $objCurrent.Name | Set-Contact -LastName $objIntermediate.sn
		}
		if($objCurrent.displayName -ne $objIntermediate.displayName){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Updated Display Name for $($objCurrent.Name) OLD VALUE: $($objCurrent.displayName) NEW VALUE: $($objIntermediate.displayName)"
			Get-MailContact $objCurrent.Name | Set-Contact -DisplayName $objIntermediate.displayName
		}
		if(($objCurrent.department -ne $objIntermediate.department) -and !([string]::IsNullOrEmpty($objCurrent.department))){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Updated Department for $($objCurrent.Name) OLD VALUE: $($objCurrent.department) NEW VALUE: $($objIntermediate.department)"
			Get-MailContact $objCurrent.Name | Set-Contact -Department $objIntermediate.department
		}
		if(($objCurrent.telephoneNumber -ne $objIntermediate.telephoneNumber) -and !([string]::IsNullOrEmpty($objCurrent.telephoneNumber))){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Updated Telephone Number for $($objCurrent.Name) OLD VALUE: $($objCurrent.telephoneNumber) NEW VALUE: $($objIntermediate.telephoneNumber)"
			Get-MailContact $objCurrent.Name | Set-Contact -Phone $objIntermediate.telephoneNumber
		}
		if(($objCurrent.title -ne $objIntermediate.title) -and !([string]::IsNullOrEmpty($objCurrent.title))){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Updated Title for $($objCurrent.Name) OLD VALUE: $($objCurrent.title) NEW VALUE: $($objIntermediate.title)"
			Get-MailContact $objCurrent.Name | Set-Contact -Title $objIntermediate.title
		}
		if(($objCurrent.CustomAttribute1 -ne $strCustomAttribute1) -and !([string]::IsNullOrEmpty($objCurrent.CustomAttribute1))){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: Custom Attribute for $($objCurrent.Name) OLD VALUE: $($objCurrent.CustomAttribute1) NEW VALUE: $strCustomAttribute1"
			Get-MailContact $objCurrent.Name | Set-MailContact -CustomAttribute1 $strCustomAttribute1
		}		
		if($objCurrent.mail -ne $objIntermediate.mail){
			# Check if the proxy address does not currently exists/is not assigned to another user
			[string] $newMailAddress = $objIntermediate.mail
			$mailExists = Get-ADObject -Properties mail, proxyAddresses -Filter {mail -eq $newMailAddress -or proxyAddresses -eq $newMailAddress}
			if($mailExists.Count -eq 0){
				Add-Content $strAuditLog "$(Get-Date) - INFO :: Email Address for $($objCurrent.Name) OLD VALUE: $($objCurrent.mail) NEW VALUE: $($objIntermediate.mail)"
				Get-MailContact $objCurrent.Name | Set-MailContact -ExternalEmailAddress $objIntermediate.mail
			} else {
				Add-Content $strAuditLog "$(Get-Date) - WARNING :: An object with the email address '$($objCurrent.mail)' already exists. The Object named '$($objCurrent.Name)' has not been updated."
				Write-Warning "An object with the email address '$($objCurrent.mail)' already exists. The Object named '$($objCurrent.Name)' has not been updated."
			}
		}
	}
}
# Next we need to add new contacts that currently do not exist within the organisation
foreach($objCurrent in $objCSVImport){
	# Search for the object in the Active Directory collection
	$objIntermediate = $colContactObjects | ?{$_.name -eq $objCurrent.name}
	if($objIntermediate.Count -eq 0){
		# Check if the proxy address does not currently exists/is not assigned to another user
		[string] $newMailAddress = $objCurrent.mail
		$mailExists = Get-ADObject -Properties mail, proxyAddresses -Filter {mail -eq $newMailAddress -or proxyAddresses -eq $newMailAddress}
		if($mailExists.Count -eq 0){
			Add-Content $strAuditLog "$(Get-Date) - INFO :: A new contact has been added $($objCurrent.Name)"
			New-MailContact -FirstName "$($objCurrent.givenName)" -LastName "$($objCurrent.sn)" -Name "$($objCurrent.Name)" -DisplayName "$($objCurrent.displayName)" -ExternalEmailAddress "$($objCurrent.mail)" -OrganizationalUnit "$($strTargetOUDN)" > $nul
			Get-MailContact $objCurrent.Name | Set-Contact -Department $objCurrent.department -Phone $objCurrent.telephoneNumber -Title $objCurrent.title -Company $strCompany
			if(!([string]::IsNullOrEmpty($strCustomAttribute1))){
				Get-MailContact $objCurrent.Name | Set-MailContact -CustomAttribute1 $strCustomAttribute1
			}
		}
		else {
			Add-Content $strAuditLog "$(Get-Date) - WARNING :: An object with the email address '$($objCurrent.mail)' already exists. The Object named '$($objCurrent.Name)' has not been added."
			Write-Warning "An object with the email address '$($objCurrent.mail)' already exists. The Object named '$($objCurrent.Name)' has not been added."
		}
    }
}
Remove-PSSession $exchangeSession
