# Documentation by Carl Webster

## Docu-XD7_V2.ps1
	Creates an inventory of a Citrix XenDesktop 7.8 - 2003 Site (from CVAD 2006 Docu-CVAD7_V3.ps1 must be used) using Microsoft PowerShell, Word, 
	plain text, or HTML.
	
	This Script requires at least PowerShell version 3 but runs best in version 5.

	Word is NOT needed to run the script. This script will output in Text and HTML.
	
	You do NOT have to run this script on a Controller. This script was developed and run 
	from a Windows 10 VM.
	
	You can run this script remotely using the –AdminAddress (AA) parameter.
	
	This script supports versions of XenApp/XenDesktop starting with 7.8.
	
	NOTE: The account used to run this script must have at least Read access to the SQL 
	Server(s) that hold(s) the Citrix Site, Monitoring, and Logging databases.
	
	By default, only gives summary information for:
		Administrators
		App-V Publishing
		AppDisks
		AppDNA
		Application Groups
		Applications
		Controllers
		Delivery Groups
		Hosting
		Logging
		Machine Catalogs
		Policies
		StoreFront
		Zones

	The Summary information is what is shown in the top half of Citrix Studio for:
		Machine Catalogs
		AppDisks
		Delivery Groups
		Applications
		Policies
		Logging
		Controllers
		Administrators
		Hosting
		StoreFront

	Using the MachineCatalogs parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.
	
	Using the DeliveryGroups parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.

	Using both the MachineCatalogs and DeliveryGroups parameters can cause the report to 
	take an extremely long time to complete and generate an exceptionally long report.
	
	Using BrokerRegistryKeys requires the script is run elevated.

	Creates an output file named after the XenDesktop 7.8+ Site.
	
	Word and PDF Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

## Docu-CVAD7_V3.ps1
	Creates an inventory of a Citrix Virtual Apps and Desktops (CVAD) 2006 or later Site 
	using Microsoft PowerShell, Word, plain text, or HTML.
	
	This Script requires at least PowerShell version 5.

	Default output is now HTML.
	
	You do NOT have to run this script on a Controller. This script was developed and run 
	from a Windows 10 VM.
	
	You can run this script remotely using the –AdminAddress (AA) parameter.
	
	This script supports versions of CVAD starting with 2006.
	
	NOTE: The account used to run this script must have at least Read access to the SQL 
	Server(s) that hold(s) the Citrix Site, Monitoring, and Logging databases.
	
	By default, only gives summary information for:
		Administrators
		App-V Publishing
		Application Groups
		Applications
		Controllers
		Delivery Groups
		Hosting
		Logging
		Machine Catalogs
		Policies
		StoreFront
		Zones

	The Summary information is what is shown in the top half of Citrix Studio for:
		Machine Catalogs
		Delivery Groups
		Applications
		Policies
		Logging
		Controllers
		Administrators
		Hosting
		StoreFront

	Using the MachineCatalogs parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.
	
	Using the DeliveryGroups parameter can cause the report to take a very long time to 
	complete and can generate an extremely long report.

	Using both the MachineCatalogs and DeliveryGroups parameters can cause the report to 
	take an extremely long time to complete and generate an exceptionally long report.
	
	Using BrokerRegistryKeys requires the script is run elevated.

	Creates an output file named after the CVAD Site.
	
	Word and PDF Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

## Docu-FAS.ps1
	Creates an inventory of Citrix Federated Authentication Service using Microsoft 
	PowerShell, Word, plain text, or HTML.
	
	This Script requires at least PowerShell version 3.
	
	This script requires an elevated PowerShell session.

	Word is NOT needed to run the script. This script will output in Text and HTML.
	The default output format is HTML.
	
	Creates an output file named CitrixFASInventory.<fileextension>.
	
	Word and PDF Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish-add	
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

## Docu-ADC.ps1
	Creates a complete inventory of a Citrix ADC configuration using Microsoft Word and PowerShell.
	Creates a Word document named after the Citrix ADC Configuration.
	Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	Script requires at least PowerShell version 3 but runs best in version 5.

## Docu-GPO.ps1
	Creates a Backup and HTML and XML Reports for all Group Policies in the current Active Directory domain.

	This Script requires at least PowerShell version 3 but runs best in version 5.

	This script requires at least one domain controller running Windows Server 2008 R2.
	
	This script outputs Text, XML and HTML files.
	
	You do NOT have to run this script on a domain controller, and it is best if you didn't.

	This script was developed and run from a Windows 10 domain-joined VM.

	This script requires Domain Admin rights and an elevated PowerShell session.
	
	To run the script from a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		http://www.microsoft.com/en-us/download/details.aspx?id=7887
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520

## Docu-DHCP.ps1
	Creates a complete inventory of a Microsoft 2012+ DHCP server using Microsoft 
	PowerShell, Word, plain text, or HTML.
	
	Creates a Word or PDF document, text or HTML file named either:
		DHCP Inventory Report for Server <DHCPServerName> for the Domain <domain>.HTML 
		(or .DOCX or .PDF or .TXT).
		DHCP Inventory Report for All DHCP Servers for the Domain <domain>.HTML (or .DOCX 
		or .PDF or .TXT).

	Version 2.0 changes the default output report from Word to HTML.
	
	The script requires at least PowerShell version 4 but runs best in version 5.
	
	Word is NOT needed to run the script. This script outputs in Text and HTML.

	You do NOT have to run this script on a DHCP server. This script was developed 
	and run from a Windows 10 VM.

	Requires the DHCPServer module.
	
	The script can run on a DHCP server or a Windows 8.x or Windows 10 computer with RSAT installed.
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
	For Windows Server 2003, 2008, and 2008 R2, use the following to export and import the 
	DHCP data:
		Export from the 2003, 2008, or 2008 R2 server:
			netsh dhcp server export C:\DHCPExport.txt all
			
			Copy the C:\DHCPExport.txt file to the 2012+ server.
			
		Import on the 2012+ server:
			netsh dhcp server import c:\DHCPExport.txt all
			
		The script can now be run on the 2012+ DHCP server to document the older DHCP 
		information.

	For Windows Server 2008 and Server 2008 R2, the 2012+ DHCP Server PowerShell cmdlets 
	can be used for export and import.
		From the 2012+ DHCP server:
			Export-DhcpServer -ComputerName 2008R2Server.domain.tld -Leases -File 
			C:\DHCPExport.xml 
			
			Import-DhcpServer -ComputerName 2012Server.domain.tld -Leases -File 
			C:\DHCPExport.xml -BackupPath C:\dhcp\backup\ 
			
			Note: The c:\dhcp\backup path must exist before running the 
			Import-DhcpServer cmdlet.
	
	Using netsh is much faster than using the PowerShell export and import cmdlets.
	
	Processing of IPv4 Multicast Scopes is only available with Server 2012 R2 DHCP.
	
	Word and PDF Documents include a Cover Page, Table of Contents, and Footer.
	
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
