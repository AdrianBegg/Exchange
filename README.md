# Exchange
Scripts for Exchange related stuff

This repository has a number of scripts which may be handy for Exchange 2013/2016 administration. Below is brief description of each and thier purpose.

Exchange-TransportLogAnalyser.ps1 - The purpose of this script is to generate a summary report from the SMTP Transport logs of hosts that have connected to the connector. This is intended to be used when migrating/decommission of an Exchange server or connector to determine what (e.g.. Rouge MFD's or Applications where someone just put in an IP address 20 years ago) is still connecting. Outputs a simple CSV file.
