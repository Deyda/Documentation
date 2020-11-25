# Documentation by Carl Webster

## Docu-XD7_V2.ps1
Creates an inventory of a Citrix XenDesktop 7.8+ Site using Microsoft PowerShell, Word, 
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
