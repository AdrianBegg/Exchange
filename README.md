# Exchange
Scripts for Exchange related stuff

This repository has a number of scripts which may be handy for Exchange 2013/2016 administration. Below is brief description of each and thier purpose.

Exchange-ReceiveConnectorLogAnalyser.ps1 - The purpose of this script is to generate a summary report from the SMTP Transport logs of hosts that have connected to the connector. This is intended to be used when migrating/decommission of an Exchange server or connector to determine what (e.g.. Rouge MFD's or Applications where someone just put in an IP address 20 years ago) is still connecting. Outputs a simple CSV file.

Import-Organisation-Contacts.ps1 - The purpose of the script is to create Mail Contacts in an organisation from user objects in a foriegn Active Directory forest. The primary use case for this is two seperate Active Directory forests with Partner Organisations where the Exchange Organisations must be kept seperate (eg. for legal reasons) however users wish to have these contacts within thier GAL.
