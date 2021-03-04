#Requires -Version 3.0
#This File is in Unicode format. Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

#region help text

<#
.SYNOPSIS
	Creates an inventory of Citrix Federated Authentication Service.
.DESCRIPTION
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
		
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	
	HTML is now the default report format.
	
	This parameter is set True if no other output format is selected.
.PARAMETER MSWord
	SaveAs DOCX file
	
	Microsoft Word is no longer the default report format.
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	
	This parameter requires Microsoft Word to be installed.
	This parameter uses Word's SaveAs PDF capability.

	This parameter is disabled by default.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date timestamp to the end of the file name.
	
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	
	Output filename will be ReportName_2020-06-01_1800.<fileextension>.
	
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER CitrixTemplatesOnly
	When processing the Microsoft certificate templates, only process the templates
	with Citrix in the template name.
	
	If you make a copy of a Citrix template and do not include "Citrix" in the
	template name, that template is not included.

	This parameter is disabled by default.
	This parameter has an alias of CTO.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field. 
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field. 
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name to use for the Cover Page. 
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field. 
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)

	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)

	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER From
	Specifies the username for the From email address.
	
	If SmtpServer or To are used, this is a required parameter.
.PARAMETER Hardware
	Use WMI to gather hardware information on Computer System, Disks, Processor, and 
	Network Interface Cards for the Certificate Authority server(s) and the FAS 
	server(s).

	This parameter may require using an account with permission to retrieve hardware 
	information (i.e. Domain Admin or Local Administrator, this includes Local 
	Administrator on the Certificate Authority server(s)).

	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.

	This parameter is disabled by default.
	This parameter has an alias of HW.
.PARAMETER LimitUserCertificates
	Use this parameter to limit the number of FAS user certificates included in the 
	report.
	
	By default, this is set to [int]::MaxValue which is 2,147,483,647.
	This parameter has an alias of LUC.
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER SmtpPort
	Specifies the SMTP port for the SmtpServer. 
	The default is 25.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report(s). 
	
	If From or To are used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	
	If SmtpServer or From are used, this is a required parameter.
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1
	
	Creates an HTML report.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -TEXT
	
	Creates a text report.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -HTML

	Creates an HTML report.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -HTML -MSWord -PDF -Text
	
	Creates four reports. One each in HTML, Microsoft Word, PDF, and plain text.
.EXAMPLE
	PS C:\PSScript .\Docu-FAS.ps1 -MSWord -CompanyName "Carl Webster 
	Consulting" -CoverPage "Mod" -UserName "Carl Webster"
	
	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\Docu-FAS.ps1 -MSWord -CN "Carl Webster Consulting" -CP 
	"Mod" -UN "Carl Webster"
	
	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		The computer running the script for the AdminAddress.
.EXAMPLE
	PS C:\PSScript .\Docu-FAS.ps1 -MSWord 
	-CompanyName "Sherlock Holmes Consulting" 
	-CoverPage Exposure 
	-UserName "Dr. Watson"
	-CompanyAddress "221B Baker Street, London, England"
	-CompanyFax "+44 1753 276600"
	-CompanyPhone "+44 1753 276200"
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.
.EXAMPLE
	PS C:\PSScript .\Docu-FAS.ps1 -MSWord 
	-CompanyName "Sherlock Holmes Consulting" 
	-CoverPage Facet 
	-UserName "Dr. Watson"
	-CompanyEmail SuperSleuth@SherlockHolmes.com

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -AddDateTime
	
	Creates an HTML file.
	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be CitrixFASInventory_2020-06-01_1800.html.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -PDF -AddDateTime
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be CitrixFASInventory_2020-06-01_1800.pdf.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -Folder \\FileServer\ShareName
	
	HTML output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -Dev -ScriptInfo -Log
	
	Creates an HTML file.
	
	Creates a text file named FASV1InventoryScriptErrors_yyyy-MM-dd_HHmm.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named FASV1InventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	FASV1DocScriptTranscript_yyyy-MM-dd_HHmm.txt.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -CitrixTemplatesOnly
	
	Creates an HTML report.
	Includes only certificate templates with "Citrix" in the name.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 -LimitUserCertificates 25
	
	Creates an HTML report.
	Includes up to the first 25 user certificates.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 
	-SmtpServer mailrelay.domain.tld
	-From Anonymous@domain.tld 
	-To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and will not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script will generate an anonymous secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 
	-SmtpServer labaddomain-com.mail.protection.outlook.com
	-UseSSL
	-From SomeEmailAddress@labaddomain.com 
	-To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 
	-SmtpServer smtp.office365.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Docu-FAS.ps1 
	-SmtpServer smtp.gmail.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.INPUTS
	None. You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script. 
	This script creates a Word, PDF, plain text, or HTML document.
.NOTES
	NAME: Docu-FAS.ps1
	VERSION: 1.11
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: May 9, 2020
#>

#endregion

#region script parameters
#thanks to Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("CTO")]
	[Switch]$CitrixTemplatesOnly=$False,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[Alias("HW")]
	[Switch]$Hardware=$False,

	[parameter(Mandatory=$False)] 
	[Alias("LUC")]
	[Int]$LimitUserCertificates=[int]::MaxValue,

	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(Mandatory=$False)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False

	)
#endregion

#region script change log	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on March 31, 2019
#
#Version 1.11 9-May-2020
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Add Receive Side Scaling setting to Function OutputNICItem
#	Change color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Update Function SendEmail to handle anonymous unauthenticated email
#	Update Function SetWordCellFormat to change parameter $BackgroundColor to [int]
#	Update Functions GetComputerWMIInfo and OutputNicInfo to fix two bugs in NIC Power Management settings
#	Update Help Text
#
#Version 1.10 10-Feb-2020
#	Added new section for Local Administration Policy
#	Added new section for Private Key Pool Info
#	Added to FAS Server Information, the Capabilities property
#	Fixed several alignment issues in the Text output option
#	Fixed Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehÃ¥llsfÃ¶rteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehÃ¥llsfÃ¶rteckn2'; Break }
#
#Version 1.01 18-May-2019
#	Fix some typos in the help text and remove some unneeded comments
#
# Version 1.0 released to the community on May 13, 2019
#
#endregion

#region initial variable testing and setup
Set-StrictMode -Version Latest

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$Script:emailCredentials = $Null

function wv
{
	$s = $args -join ''
	Write-Verbose $s
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "
			`n`n
			`t`t
			Folder $Folder is a file, not a folder.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			AbortScript
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
		`t`t
		Folder $Folder does not exist.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\FASV1DocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdpath\FASV1InventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
#endregion

#region initialize variables for Word, HTML, and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdColorGray15 = 14277081
	[int]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	#[int]$wdAlignParagraphLeft = 0
	#[int]$wdAlignParagraphCenter = 1
	#[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	#[int]$wdCellAlignVerticalTop = 0
	#[int]$wdCellAlignVerticalCenter = 1
	#[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	#[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	#[int]$wdAdjustFirstColumn = 2
	#[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	#[int]$Indent1TabStops = 1 * $PointsPerTabStop
	#[int]$Indent2TabStops = 2 * $PointsPerTabStop
	#[int]$Indent3TabStops = 3 * $PointsPerTabStop
	#[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155
	[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	#[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
    $Script:htmlredmask       = "#FF0000" 4>$Null
    $Script:htmlcyanmask      = "#00FFFF" 4>$Null
    $Script:htmlbluemask      = "#0000FF" 4>$Null
    $Script:htmldarkbluemask  = "#0000A0" 4>$Null
    $Script:htmllightbluemask = "#ADD8E6" 4>$Null
    $Script:htmlpurplemask    = "#800080" 4>$Null
    $Script:htmlyellowmask    = "#FFFF00" 4>$Null
    $Script:htmllimemask      = "#00FF00" 4>$Null
    $Script:htmlmagentamask   = "#FF00FF" 4>$Null
    $Script:htmlwhitemask     = "#FFFFFF" 4>$Null
    $Script:htmlsilvermask    = "#C0C0C0" 4>$Null
    $Script:htmlgraymask      = "#808080" 4>$Null
    $Script:htmlblackmask     = "#000000" 4>$Null
    $Script:htmlorangemask    = "#FFA500" 4>$Null
    $Script:htmlmaroonmask    = "#800000" 4>$Null
    $Script:htmlgreenmask     = "#008000" 4>$Null
    $Script:htmlolivemask     = "#808000" 4>$Null

    $Script:htmlbold        = 1 4>$Null
    $Script:htmlitalics     = 2 4>$Null
    $Script:htmlred         = 4 4>$Null
    $Script:htmlcyan        = 8 4>$Null
    $Script:htmlblue        = 16 4>$Null
    $Script:htmldarkblue    = 32 4>$Null
    $Script:htmllightblue   = 64 4>$Null
    $Script:htmlpurple      = 128 4>$Null
    $Script:htmlyellow      = 256 4>$Null
    $Script:htmllime        = 512 4>$Null
    $Script:htmlmagenta     = 1024 4>$Null
    $Script:htmlwhite       = 2048 4>$Null
    $Script:htmlsilver      = 4096 4>$Null
    $Script:htmlgray        = 8192 4>$Null
    $Script:htmlolive       = 16384 4>$Null
    $Script:htmlorange      = 32768 4>$Null
    $Script:htmlmaroon      = 65536 4>$Null
    $Script:htmlgreen       = 131072 4>$Null
	$Script:htmlblack       = 262144 4>$Null

	$Script:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$Script:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region code for hardware data
Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	If($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteHTMLLine 4 0 "General Computer"
	}
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select-Object Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null
		[string]$ComputerOS = (Get-WmiObject -class Win32_OperatingSystem -computername $RemoteComputerName -EA 0).Caption

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item $ComputerOS
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" -option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" -Option $htmlBold
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	If($Text)
	{
		Line 1 "Drive(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Drive(s)"
	}

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$drives = $Results | Select-Object caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" -Option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" -Option $htmlBold
		}
	}
	
	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	If($Text)
	{
		Line 1 "Processor(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Processor(s)"
	}

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select-Object availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" -Option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" -Option $htmlBold
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	If($Text)
	{
		Line 1 "Network Interface(s)"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					OutputNicItem $Nic $ThisNic $RemoteComputerName
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" -Option $htmlBold
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" -Option $htmlBold
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" -Option $htmlBold
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" -Option $htmlBold
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" -Option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" -Option $htmlBold
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS)
	
	If($MSWord -or $PDF)
	{
		$ItemInformation = New-Object System.Collections.ArrayList
		$ItemInformation.Add(@{ Data = "Manufacturer"; Value = $Item.manufacturer; }) > $Null
		$ItemInformation.Add(@{ Data = "Model"; Value = $Item.model; }) > $Null
		$ItemInformation.Add(@{ Data = "Domain"; Value = $Item.domain; }) > $Null
		$ItemInformation.Add(@{ Data = "Operating System"; Value = $OS; }) > $Null
		$ItemInformation.Add(@{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }) > $Null
		$ItemInformation.Add(@{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }) > $Null
		$ItemInformation.Add(@{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }) > $Null
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 2 "Model`t`t`t`t: " $Item.model
		Line 2 "Domain`t`t`t`t: " $Item.domain
		Line 2 "Operating System`t`t: " $OS
		Line 2 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlBold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlBold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlBold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlBold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlBold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlBold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"; Break}
		1	{$xDriveType = "No Root Directory"; Break}
		2	{$xDriveType = "Removable Disk"; Break}
		3	{$xDriveType = "Local Disk"; Break}
		4	{$xDriveType = "Network Drive"; Break}
		5	{$xDriveType = "Compact Disc"; Break}
		6	{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		$DriveInformation = New-Object System.Collections.ArrayList
		$DriveInformation.Add(@{ Data = "Caption"; Value = $Drive.caption; }) > $Null
		$DriveInformation.Add(@{ Data = "Size"; Value = "$($drive.drivesize) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation.Add(@{ Data = "File System"; Value = $Drive.filesystem; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation.Add(@{ Data = "Volume Name"; Value = $Drive.volumename; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation.Add(@{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation.Add(@{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Drive Type"; Value = $xDriveType; }) > $Null
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	If($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlBold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlBold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlBold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlBold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlBold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlBold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlBold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlBold),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break}
		2	{$xAvailability = "Unknown"; Break}
		3	{$xAvailability = "Running or Full Power"; Break}
		4	{$xAvailability = "Warning"; Break}
		5	{$xAvailability = "In Test"; Break}
		6	{$xAvailability = "Not Applicable"; Break}
		7	{$xAvailability = "Power Off"; Break}
		8	{$xAvailability = "Off Line"; Break}
		9	{$xAvailability = "Off Duty"; Break}
		10	{$xAvailability = "Degraded"; Break}
		11	{$xAvailability = "Not Installed"; Break}
		12	{$xAvailability = "Install Error"; Break}
		13	{$xAvailability = "Power Save - Unknown"; Break}
		14	{$xAvailability = "Power Save - Low Power Mode"; Break}
		15	{$xAvailability = "Power Save - Standby"; Break}
		16	{$xAvailability = "Power Cycle"; Break}
		17	{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	If($MSWORD -or $PDF)
	{
		$ProcessorInformation = New-Object System.Collections.ArrayList
		$ProcessorInformation.Add(@{ Data = "Name"; Value = $Processor.name; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Description"; Value = $Processor.description; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }) > $Null
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }) > $Null
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }) > $Null
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Cores"; Value = $Processor.numberofcores; }) > $Null
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }) > $Null
		}
		$ProcessorInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Name`t`t`t`t: " $processor.name
		Line 2 "Description`t`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t`t: " $xAvailability
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlBold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlBold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlBold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlBold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlBold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlBold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlBold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlBold),$xAvailability,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic, [string]$RemoteComputerName)
	
	$powerMgmt = Get-WmiObject -computername $RemoteComputerName MSPower_DeviceEnable -Namespace root\wmi | Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

	If($? -and $Null -ne $powerMgmt)
	{
		If($powerMgmt.Enable -eq $True)
		{
			$PowerSaving = "Enabled"
		}
		Else
		{
			$PowerSaving = "Disabled"
		}
	}
	Else
	{
        $PowerSaving = "N/A"
	}
	
	$xAvailability = ""
	Switch ($ThisNic.availability)
	{
		1		{$xAvailability = "Other"; Break}
		2		{$xAvailability = "Unknown"; Break}
		3		{$xAvailability = "Running or Full Power"; Break}
		4		{$xAvailability = "Warning"; Break}
		5		{$xAvailability = "In Test"; Break}
		6		{$xAvailability = "Not Applicable"; Break}
		7		{$xAvailability = "Power Off"; Break}
		8		{$xAvailability = "Off Line"; Break}
		9		{$xAvailability = "Off Duty"; Break}
		10		{$xAvailability = "Degraded"; Break}
		11		{$xAvailability = "Not Installed"; Break}
		12		{$xAvailability = "Install Error"; Break}
		13		{$xAvailability = "Power Save - Unknown"; Break}
		14		{$xAvailability = "Power Save - Low Power Mode"; Break}
		15		{$xAvailability = "Power Save - Standby"; Break}
		16		{$xAvailability = "Power Cycle"; Break}
		17		{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	#attempt to get Receive Side Scaling setting
	$RSSEnabled = "N/A"
	Try
	{
		#https://ios.developreference.com/article/10085450/How+do+I+enable+VRSS+(Virtual+Receive+Side+Scaling)+for+a+Windows+VM+without+relying+on+Enable-NetAdapterRSS%3F
		$RSSEnabled = (Get-WmiObject -ComputerName $RemoteComputerName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0).Enabled

		If($RSSEnabled)
		{
			$RSSEnabled = "Enabled"
		}
		ELse
		{
			$RSSEnabled = "Disabled"
		}
	}
	
	Catch
	{
		$RSSEnabled = "Not available on $Script:RunningOS"
	}

	$xIPAddress = New-Object System.Collections.ArrayList
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress.Add("$($IPAddress)") > $Null
	}

	$xIPSubnet = New-Object System.Collections.ArrayList
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet.Add("$($IPSubnet)") > $Null
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder.Add("$($DNSDomain)") > $Null
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder.Add("$($DNSServer)") > $Null
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		$NicInformation = New-Object System.Collections.ArrayList
		$NicInformation.Add(@{ Data = "Name"; Value = $ThisNic.Name; }) > $Null
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation.Add(@{ Data = "Description"; Value = $Nic.description; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }) > $Null
		If(validObject $Nic Manufacturer)
		{
			$NicInformation.Add(@{ Data = "Manufacturer"; Value = $Nic.manufacturer; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$NicInformation.Add(@{ Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }) > $Null
		$NicInformation.Add(@{ Data = "Receive Side Scaling"; Value = $RSSEnabled; }) > $Null
		$NicInformation.Add(@{ Data = "Physical Address"; Value = $Nic.macaddress; }) > $Null
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress[0]; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = "IP Address"; Value = $tmp; }) > $Null
					$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }) > $Null
				}
			}
		}
		Else
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet; }) > $Null
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation.Add(@{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation.Add(@{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }) > $Null
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }) > $Null
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }) > $Null
		$NicInformation.Add(@{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }) > $Null
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation.Add(@{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation.Add(@{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation.Add(@{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation.Add(@{ Data = "Scope ID"; Value = $Nic.winsscopeid; }) > $Null
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		If(validObject $Nic Manufacturer)
		{
			Line 2 "Manufacturer`t`t: " $Nic.manufacturer
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Allow computer to turn "
		Line 2 "off device to save power: " $PowerSaving
		Line 2 "Receive Side Scaling`t: " $RSSEnabled
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t: " $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t: " $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlBold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlBold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlBold),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlBold),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlBold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($htmlsilver -bor $htmlBold),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Receive Side Scaling',($htmlsilver -bor $htmlbold),$RSSEnabled,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlBold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlBold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlBold),$Nic.Defaultipgateway[0],$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlBold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlBold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlBold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlBold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlBold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlBold),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlBold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlBold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlBold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlBold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlBold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlBold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlBold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlBold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlBold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automÃ¡tica 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automÃ¡tica 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
#			'fr-'	{ 'Sommaire Automatique 2'; Break }
			'fr-'	{ 'Table automatiqueÂ 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'SumÃ¡rio AutomÃ¡tico 2'; Break }
			# fix in 1.90 thanks to Johan Kallio 'sv-'	{ 'Automatisk innehÃ¥llsfÃ¶rteckning2'; Break }
			'sv-'	{ 'Automatisk innehÃ¥llsfÃ¶rteckn2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_}	{$CultureCode = "ca-"}
		{$DanishArray -contains $_}		{$CultureCode = "da-"}
		{$DutchArray -contains $_}		{$CultureCode = "nl-"}
		{$EnglishArray -contains $_}	{$CultureCode = "en-"}
		{$FinnishArray -contains $_}	{$CultureCode = "fi-"}
		{$FrenchArray -contains $_}		{$CultureCode = "fr-"}
		{$GermanArray -contains $_}		{$CultureCode = "de-"}
		{$NorwegianArray -contains $_}	{$CultureCode = "nb-"}
		{$PortugueseArray -contains $_}	{$CultureCode = "pt-"}
		{$SpanishArray -contains $_}	{$CultureCode = "es-"}
		{$SwedishArray -contains $_}	{$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "IÃ³ (clar)", "IÃ³ (fosc)", "LÃ­nia lateral",
					"Moviment", "QuadrÃ­cula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "SemÃ for", "VisualitzaciÃ³ principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "IÃ³ (clar)", "IÃ³ (fosc)", "LÃ­nia lateral",
					"Moviment", "QuadrÃ­cula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "SemÃ for", "VisualitzaciÃ³", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "DiplomÃ tic", "ExposiciÃ³",
					"LÃ­nia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "QuadrÃ­cula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevÃ¦gElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mÃ¸rk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mÃ¸rk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevÃ¦gElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mÃ¸rk)", "Ion (mÃ¸rk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevÃ¦gElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "GÃ¥de",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"NÃ¥lestribet", "Ã…rlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"GebÃ¤ndert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "RÃ¼ckblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "RÃ¼ckblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "JÃ¤hrlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "CuadrÃ­cula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "LÃ­nea lateral", "Movimiento", "Retrospectiva", 
					"SemÃ¡foro", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "SemÃ¡foro", "Retrospectiva", "CuadrÃ­cula",
					"Movimiento", "Cortar (oscuro)", "LÃ­nea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "CuadrÃ­cula", "CubÃ­culos", "ExposiciÃ³n", "LÃ­nea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periÃ³dico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "VaihtuvavÃ¤rinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "VaihtuvavÃ¤rinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Ã€ bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "IntÃ©grale", "Ion (clair)", "Ion (foncÃ©)", 
					"Lignes latÃ©rales", "Quadrillage", "RÃ©trospective", "Secteur (clair)", 
					"Secteur (foncÃ©)", "SÃ©maphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "AustÃ¨re", "Austin", 
					"Blocs empilÃ©s", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latÃ©rale", "Moderne", 
					"MosaÃ¯ques", "Mots croisÃ©s", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mÃ¸rk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mÃ¸rk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Ã…rlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("AnimaÃ§Ã£o", "Austin", "Em Tiras", "ExibiÃ§Ã£o Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Ãon (Claro)", "Ãon (Escuro)", "Linha Lateral",
					"Retrospectiva", "SemÃ¡foro")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "AnimaÃ§Ã£o", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "ExposiÃ§Ã£o", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeÃ§a", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mÃ¶rkt)", "Knippe", "RutnÃ¤t", "RÃ¶rElse", "Sektor (ljus)", "Sektor (mÃ¶rk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Ã…terblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("AlfabetmÃ¶nster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "RutnÃ¤t",
					"RÃ¶rElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Ã…rligt",
					"Ã-vergÃ¥ende")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		AbortScript
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	#fixed by MBS
	[bool]$wordrunning = $null -ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		AbortScript
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		Write-Error "
		`n`n
		`t`t
		The Word object could not be created. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		Write-Error "
		`n`n
		`t`t
		Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		Write-Error "
		`n`n
		`t`t
		Microsoft Word 2007 is no longer supported.
		`n`n
		`t`t
		Script will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
		`t`t
		The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
	Else
	{
		Write-Error "
		`n`n
		`t`t
		You are running an untested or unsupported version of Microsoft Word.
		`n`n
		`t`t
		Script will end.
		`n`n
		`t`t
		Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "
			`n`n
			`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "
			`n
			`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "
			`n
			`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			`n`n"
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "LÃ­nia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "LÃ­nea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latÃ©rales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latÃ©rale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
		`t`t
		For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		Write-Error "
		`n`n
		`t`t
		An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		Write-Error "
		`n`n
		`t`t
		An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object{$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue2
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey = $Reg.OpenSubKey($path2)
		If ($RegKey)
		{
			$Results = $RegKey.GetValue($name)

			If($Null -ne $Results)
			{
				Return $Results
			}
			Else
			{
				Return $Null
			}
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		$null = $Script:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		$null = $Script:Output.Append( $name + $value )
	}
	Else
	{
		$null = $Script:Output.AppendLine( $name + $value )
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML. They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used. Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

# to suppress $crlf in HTML documents, replace this with '' (empty string)
# but this was added to make the HTML readable
$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
# re-implemented/re-based by Michael B. Smith
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	#FIXME - long story short, this function was wrong and had been wrong for a long time. 
	## The function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites! If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		## $HTMLBody = '<p></p>'
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold
		if( $ital ) { $null = $sb.Append( '<i>' ) }
		if( $bold ) { $null = $sb.Append( '<b>' ) } 

		switch( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		$null = $sb.Append( $HTMLOpen )

		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		if( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		else                    { $null = $sb.Append( $HTMLClose ) }

		if( $ital ) { $null = $sb.Append( '</i>' ) }
		if( $bold ) { $null = $sb.Append( '</b>' ) } 

		if( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)

	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	if( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
		$rowCount = $rowInfo.Length
	}

	for( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		if( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
		}

		$subRowLength = $subRow.Length
		for( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = if( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } else { 0 }

			$text   = if( $item ) { $item.ToString() } else { '' }
			$format = if( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $Script:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics

			if( $null -eq $fixedInfo -or $fixedInfo.Length -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}
			else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
			}

			if( $bold ) { $null = $sb.Append( '<b>' ) }
			if( $ital ) { $null = $sb.Append( '<i>' ) }

			if( $text -eq ' ' -or $text.length -eq 0)
			{
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			else
			{
				for ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					if( $text[ $inx ] -eq ' ' )
					{
						$null = $sb.Append( '&nbsp;' )
					}
					else
					{
						break
					}
				}
				$null = $sb.Append( $text )
			}

			if( $bold ) { $null = $sb.Append( '</b>' ) }
			if( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
		}

		$null = $sb.AppendLine( '</tr>' )
	}

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith
#***********************************************************************************************************

<#
.Synopsis
	Format table for a HTML output document.
.DESCRIPTION
	This function formats a table for HTML from multiple arrays of strings.
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). Otherwise the table will be generated
	with a border (border = '1').
.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a separate array containing column headers
	(columnArray is not specified). Set this parameter equal to the number of columns in the table.
.PARAMETER rowArray
	This parameter contains the row data array for the table.
.PARAMETER columnArray
	This parameter contains column header data for the table.
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")
.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file. All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column. You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column. Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data. Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array. If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics. For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below. As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('FileName',$htmlsb,$Script:FileName,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -ne $rowArray )
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	if( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		for( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $Script:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			if( $null -eq $fixedWidth -or $fixedWidth.Length -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			if( $bold ) { $HTMLBody += '<b>' }
			if( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			if( $array )
			{
				if( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				else
				{
					for( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						if( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			if( $bold ) { $HTMLBody += '</b>' }
			if( $ital ) { $HTMLBody += '</i>' }
		}

		$HTMLBody += '</font></td>'
		$HTMLBody += $crlf
	}

	$HTMLBody += '</tr>' + $crlf

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	If( $rowArray )
	{

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HtmlFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HtmlFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HtmlFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:HtmlFileName -Force -InputObject $HTMLHead 4>$Null
}#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -eq $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$True, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $True; }
				If($Italic) { $Cell.Range.Font.Italic = $True; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$True, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$True, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += Get-PSSnapin | ForEach-Object {$_.name}
	$registeredSnapins += Get-PSSnapin -Registered | ForEach-Object {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Add-PSSnapin -Name $snapin -EA 0 *>$Null
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | ForEach-Object {Write-Warning "($_)"}
		Return $False
	}
	Else
	{
		Return $True
	}
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	Remove-Variable -Name word -Scope Script 4>$Null
	Remove-Variable -Name Doc  -Scope Script 4>$Null
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = (Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}
	If($null -ne $wordprocess -and $wordprocess.Id -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess.Id)"
		Stop-Process $wordprocess.Id -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date): Setting up Text"
	[System.Text.StringBuilder] $Script:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date): Saving Text file"
	Write-Output $Script:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date): Saving HTML file"
	Out-File -FilePath $Script:HtmlFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function ProcessDocumentOutput
{
	Param([string] $Condition)
	
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	If($Condition -eq "Regular")
	{
		$GotFile = $False

		If($MSWord)
		{
			If(Test-Path "$($Script:WordFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:WordFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:WordFileName)"
				Write-Error "Unable to save the output file, $($Script:WordFileName)"
			}
		}
		If($PDF)
		{
			If(Test-Path "$($Script:PDFFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:PDFFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:PDFFileName)"
				Write-Error "Unable to save the output file, $($Script:PDFFileName)"
			}
		}
		If($Text)
		{
			If(Test-Path "$($Script:TextFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:TextFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:TextFileName)"
				Write-Error "Unable to save the output file, $($Script:TextFileName)"
			}
		}
		If($HTML)
		{
			If(Test-Path "$($Script:HTMLFileName)")
			{
				Write-Verbose "$(Get-Date): $($Script:HTMLFileName) is ready for use"
				$GotFile = $True
			}
			Else
			{
				Write-Warning "$(Get-Date): Unable to save the output file, $($Script:HTMLFileName)"
				Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
			}
		}
		
		#email output file if requested
		If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
		{
			If($MSWord)
			{
				$emailAttachment = $Script:WordFileName
				SendEmail $emailAttachment
			}
			If($PDF)
			{
				$emailAttachment = $Script:PDFFileName
				SendEmail $emailAttachment
			}
			If($Text)
			{
				$emailAttachment = $Script:TextFileName
				SendEmail $emailAttachment
			}
			If($HTML)
			{
				$emailAttachment = $Script:HTMLFileName
				SendEmail $emailAttachment
			}
		}
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Add DateTime         : $($AddDateTime)"
	Write-Verbose "$(Get-Date): Citrix Templates Only: $($CitrixTemplatesOnly)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name         : $($Script:CoName)"
		Write-Verbose "$(Get-Date): Company Address      : $($CompanyAddress)"
		Write-Verbose "$(Get-Date): Company Email        : $($CompanyEmail)"
		Write-Verbose "$(Get-Date): Company Fax          : $($CompanyFax)"
		Write-Verbose "$(Get-Date): Company Phone        : $($CompanyPhone)"
		Write-Verbose "$(Get-Date): Cover Page           : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): Dev                  : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile         : $($Script:DevErrorFile)"
	}
	If($MSWord)
	{
		Write-Verbose "$(Get-Date): Word FileName        : $($Script:WordFileName)"
	}
	If($HTML)
	{
		Write-Verbose "$(Get-Date): HTML FileName        : $($Script:HtmlFileName)"
	} 
	If($PDF)
	{
		Write-Verbose "$(Get-Date): PDF FileName         : $($Script:PDFFileName)"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date): Text FileName        : $($Script:TextFileName)"
	}
	Write-Verbose "$(Get-Date): Folder               : $($Folder)"
	Write-Verbose "$(Get-Date): From                 : $($From)"
	Write-Verbose "$(Get-Date): Hardware Inventory   : $($Hardware)"
	Write-Verbose "$(Get-Date): Limit User Certs     : $($LimitUserCertificates)"
	Write-Verbose "$(Get-Date): Log                  : $($Log)"
	Write-Verbose "$(Get-Date): Save As HTML         : $($HTML)"
	Write-Verbose "$(Get-Date): Save As PDF          : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT         : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD         : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo           : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port            : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server          : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title                : $($Script:FASDisplayName)"
	Write-Verbose "$(Get-Date): To                   : $($To)"
	Write-Verbose "$(Get-Date): Use SSL              : $($UseSSL)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name            : $($UserName)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected          : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version         : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture            : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture          : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language        : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version         : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start         : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function AbortScript
{
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

# Converts a SDDL string into an object-based representation of a security
# descriptor

## base script found at 
## https://github.com/PowerShell/PowerShell/pull/8341/commits/518ab4fcaa245c4120ec550417bdf1287ba44062
## on 2019/04/24

## updated to decode ACL controlflags and for "set-strictmode -version latest" - MBS 2019/04/27
## make this work on PS v3 and PS v4 (no ::new() and fix several other version issues)
## note that $PSEdition was introduced in v5, so fake that for v3 and v4
## change AceQualifier into a more friendly translation (match icacls and dsacls)
## also decodes extended-rights under specific (tested) conditions

function ConvertFrom-Sddl-MBS 
{
    [CmdletBinding( HelpUri = "https://go.microsoft.com/fwlink/?LinkId=623636" )]
    Param
    (
        ## The string representing the security descriptor in SDDL syntax
        [Parameter( Mandatory = $true, Position = 0, ValueFromPipeline = $true )]
        [String] $Sddl,

        ## The type of rights that this SDDL string represents, if any.
        [Parameter( Mandatory = $false )]
        [ValidateSet(
            'FileSystemRights',
            'RegistryRights',
            'ActiveDirectoryRights',
            'MutexRights',
            'SemaphoreRights',
            'CryptoKeyRights',
            'EventWaitHandleRights'
        )]
        $Type
    )

    Begin
    {
        Set-StrictMode -Version Latest

        ## PSEdition was introduced in v5. Fake it for v3 and v4.
        $null = Get-Variable PSEdition -ErrorAction SilentlyContinue
        if( !$? )
        {
            $script:PSEdition = 'Desktop'
        } 

        # On CoreCLR, CryptoKeyRights and ActiveDirectoryRights are not supported.
        if( $PSEdition -eq 'Core' -and ( $Type -eq 'CryptoKeyRights' -or $Type -eq 'ActiveDirectoryRights' ) )
        {
            $errorId       = 'TypeNotSupported'
            $errorCategory = [System.Management.Automation.ErrorCategory]::InvalidArgument
        ##  TypeNotSupported isn't in Windows PowerShell - only in PS Core.
        ##  $errorMessage  = [Microsoft.PowerShell.Commands.UtilityResources]::TypeNotSupported -f $Type
            $errorMessage  = [Microsoft.PowerShell.Commands.UtilityResources]::AlgorithmTypeNotSupported -f $Type
            $exception     = New-Object -TypeName System.ArgumentException `
                                -ArgumentList $errorMessage
            $errorRecord   = New-Object -TypeName System.Management.Automation.ErrorRecord `
                                -ArgumentList $exception, $errorId, $errorCategory, $null
            $PSCmdlet.ThrowTerminatingError( $errorRecord )
        }

        ## Translates a SID into a NT Account
        function ConvertTo-NtAccount
        {
            Param
            (
                $Sid
            )

            if( $Sid )
            {
                $securityIdentifier = [System.Security.Principal.SecurityIdentifier] $Sid

                try
                {
                    $ntAccount = $securityIdentifier.Translate( [System.Security.Principal.NTAccount] ).ToString()
                }
                catch
                {
                    ## empty
                }

                $ntAccount
            }
        }

        ## Gets the access rights that apply to an access mask, preferring right types
        ## of 'Type' if specified.
        function Get-AccessRights
        {
            Param
            (
                $AccessMask,

                [String]
                $Type
            )

            if( $PSEdition -eq 'Core' )
            {
                ## All the types of access rights understood by .NET Core
                $rightTypes =
                    [Ordered] @{
                        'FileSystemRights'      = [System.Security.AccessControl.FileSystemRights]
                        'RegistryRights'        = [System.Security.AccessControl.RegistryRights]
                        'MutexRights'           = [System.Security.AccessControl.MutexRights]
                        'SemaphoreRights'       = [System.Security.AccessControl.SemaphoreRights]
                        'EventWaitHandleRights' = [System.Security.AccessControl.EventWaitHandleRights]
                    }
            }
            else
            {
                ## All the types of access rights understood by Windows .NET Framework
                $rightTypes =
                    [Ordered] @{
                        'FileSystemRights'      = [System.Security.AccessControl.FileSystemRights]
                        'RegistryRights'        = [System.Security.AccessControl.RegistryRights]
                        'ActiveDirectoryRights' = [System.DirectoryServices.ActiveDirectoryRights]
                        'MutexRights'           = [System.Security.AccessControl.MutexRights]
                        'SemaphoreRights'       = [System.Security.AccessControl.SemaphoreRights]
                        'CryptoKeyRights'       = [System.Security.AccessControl.CryptoKeyRights]
                        'EventWaitHandleRights' = [System.Security.AccessControl.EventWaitHandleRights]
                    }
            }
            $typesToExamine = $rightTypes.Values

            ## If they know the access mask represents a certain type, prefer its names
            ## (i.e.: CreateLink for the registry over CreateDirectories for the filesystem)
            if( $Type )
            {
                $typesToExamine = @( $rightTypes[ $Type ] ) + $typesToExamine
            }

            ## Stores the access types we've found that apply
            $foundAccess = @()

            ## Store the access types we've already seen, so that we don't report access
            ## flags that are essentially duplicate. Many of the access values in the different
            ## enumerations have the same value but with different names.
            $foundValues = @{}

            ## Go through the entries in the different right types, and see if they apply to the
            ## provided access mask. If they do, then add that to the result.   
            foreach( $rightType in $typesToExamine )
            {
                foreach( $accessFlag in [Enum]::GetNames( $rightType ) )
                {
                    $longKeyValue = [long] $rightType::$accessFlag
                    if( -not $foundValues.ContainsKey( $longKeyValue ) )
                    {
                        $foundValues[ $longKeyValue ] = $true
                        if( ( $AccessMask -band $longKeyValue ) -eq $longKeyValue )
                        {
                            $foundAccess += $accessFlag
                        }
                    }
                }
            }

            $foundAccess | Sort-Object
        }

        ## Convert ControlFlags into a string
        function ConvertTo-Control
        {
            Param
            (
                [Parameter( Mandatory = $true, ValueFromPipeline = $true )]
                $flags
            )

            ## https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.controlflags

            $None                                = 0
            $OwnerDefaulted                      = 1
            $GroupDefaulted                      = 2
            $DiscretionaryAclPresent             = 4
            $DiscretionaryAclDefaulted           = 8
            $SystemAclPresent                    = 16
            $SystemAclDefaulted                  = 32
            $DiscretionaryAclUntrusted           = 64
            $ServerSecurity                      = 128
            $DiscretionaryAclAutoInheritRequired = 256
            $SystemAclAutoInheritRequired        = 512
            $DiscretionaryAclAutoInherited       = 1024
            $SystemAclAutoInherited              = 2048
            $DiscretionaryAclProtected           = 4096
            $SystemAclProtected                  = 8192
            $RMControlValid                      = 16384
            $SelfRelative                        = 32768

            $names =
            @{
                $None                                = 'None'
                $OwnerDefaulted                      = 'OwnerDefaulted'
                $GroupDefaulted                      = 'GroupDefaulted'
                $DiscretionaryAclPresent             = 'DiscretionaryAclPresent'
                $DiscretionaryAclDefaulted           = 'DiscretionaryAclDefaulted'
                $SystemAclPresent                    = 'SystemAclPresent'
                $SystemAclDefaulted                  = 'SystemAclDefaulted'
                $DiscretionaryAclUntrusted           = 'DiscretionaryAclUntrusted'
                $ServerSecurity                      = 'ServerSecurity'
                $DiscretionaryAclAutoInheritRequired = 'DiscretionaryAclAutoInheritRequired'
                $SystemAclAutoInheritRequired        = 'SystemAclAutoInheritRequired'
                $DiscretionaryAclAutoInherited       = 'DiscretionaryAclAutoInherited'
                $SystemAclAutoInherited              = 'SystemAclAutoInherited'
                $DiscretionaryAclProtected           = 'DiscretionaryAclProtected'
                $SystemAclProtected                  = 'SystemAclProtected'
                $RMControlValid                      = 'RMControlValid'
                $SelfRelative                        = 'SelfRelative'
            }

            $explain =
            @{
                $None                                = "No control flags."
                $OwnerDefaulted                      = "Specifies that the owner SecurityIdentifier was obtained by a defaulting mechanism."
                $GroupDefaulted                      = "Specifies that the group SecurityIdentifier was obtained by a defaulting mechanism."
                $DiscretionaryAclPresent             = "Specifies that the DACL is not null."
                $DiscretionaryAclDefaulted           = "Specifies that the DACL was obtained by a defaulting mechanism."
                $SystemAclPresent                    = "Specifies that the SACL is not null."
                $SystemAclDefaulted                  = "Specifies that the SACL was obtained by a defaulting mechanism."
                $DiscretionaryAclUntrusted           = "Ignored."
                $ServerSecurity                      = "Ignored."
                $DiscretionaryAclAutoInheritRequired = "Ignored."
                $SystemAclAutoInheritRequired        = "Ignored."
                $DiscretionaryAclAutoInherited       = "Specifies that the Discretionary Access Control List (DACL) has been automatically inherited from the parent."
                $SystemAclAutoInherited              = "Specifies that the System Access Control List (SACL) has been automatically inherited from the parent."
                $DiscretionaryAclProtected           = "Specifies that the resource manager prevents auto-inheritance."
                $SystemAclProtected                  = "Specifies that the resource manager prevents auto-inheritance."
                $RMControlValid                      = "Specifies that the contents of the Reserved field are valid."
                $SelfRelative                        = "Specifies that the security descriptor binary representation is in the self-relative format. This flag is always set."
            }

            $controlFlags = @()
            $explanations = @()

            if( $flags -eq 0 )
            {
                ## $controlFlags += $names[ $None ]
                ## $explanations += "$( $names[ $None ] ) = $( $explain[ $None ] )"
                return $names[ $None ]
            }

            foreach( $key in $names.Keys )
            {
                if( $key -eq 0 )
                {
                    continue
                }
                if( ( $flags -band $key ) -eq $key )
                {
                    $controlFlags += $names[ $key ]
                    $explanations += "$( $names[ $key ] ) = $( $explain[ $key ] )"
                }
            }

            $res = ( $controlFlags | Sort-Object ) -join ', '
            ## if you want the explanations, they are here...
            ## $exp = ( $explanations | Sort-Object ) -join '; '

            return $res
        } ## end ConvertTo-Control

        function Get-ConfigNC
        {
            $rootDSE  = [ADSI]"LDAP://RootDSE"
            $configNC = $rootDSE.configurationNamingContext.Value
            $rootDSE  = $null

            $configNC
        }

        $script:ConfigNC = Get-ConfigNC

        function Get-ExtendedAttribute
        {
            Param
            (
                [Parameter( Mandatory = $true )]
                $guid
            )

            $searchRoot = 'LDAP://CN=Extended-Rights,' + $script:ConfigNC
            $ldapquery  = "(&(objectClass=controlAccessRight)(rightsGuid=$guid))"

            $adSearcher = [adsisearcher] $ldapquery
            $adSearcher.SearchRoot = [adsi] $searchRoot
            $results = $adSearcher.FindAll()

            if( -not $results )
            {
                ## didn't find the rightsGuid
                $adSearcher.SearchRoot = $null
                $adSearcher = $null

                return $guid
            }

            if( $results -is [Array] -and $results.Count -gt 0 )
            {
                ## this is bad
                Write-Error "rightsGuid search collision $guid (count $( $results.Count ))"
            }

            $displayName = ''
            foreach( $result in $results )
        	{
                $displayName = $result.Properties[ 'displayName' ]
                break
            }

            $results = $null
            $adSearcher.SearchRoot = $null
            $adSearcher = $null

            $displayName
        }

        ## more friendly name, to match icacls and dsacls
        function Get-AceQualifier
        {
            Param
            (
                [Parameter( Mandatory = $true )]
                $aceQual
            )

            $result = switch( $aceQual )
                {
                    'AccessAllowed'  { 'Allow' }
                    'AccessDenied'   { 'Deny'  }
                    'SystemAlarm'    { 'Alarm' }
                    'SystemAudit'    { 'Audit' }
                    Default          { $aceQual }
                }

            $result
        }

        ## Converts an ACE into a string representation
        function ConvertTo-AceString
        {
            Param
            (
                [Parameter( ValueFromPipeline = $true )]
                $Ace,

                $Type
            )

            Process
            {
                foreach($aceEntry in $Ace)
                {
                    $aceQual   = Get-AceQualifier $aceEntry.AceQualifier
                    $AceString = (ConvertTo-NtAccount $aceEntry.SecurityIdentifier) + ": " + $aceQual

                    if($aceEntry.AceFlags -ne "None")
                    {
                        $AceString += " " + $aceEntry.AceFlags
                    }

                    [bool] $foundExtended = $false

                    if( $aceEntry.AceType -eq 'AccessAllowedObject'         -or
                        $aceEntry.AceType -eq 'AccessAllowedCallbackObject' -or
                        $aceEntry.AceType -eq 'AccessDeniedObject'          -or
                        $aceEntry.AceType -eq 'AccessDeniedCallbackObject'
                    )
                    {
                        ## see 
                        ## https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.acetype
                        ## for a (bad) explanation on how these work
                        if( $aceEntry.ObjectAceFlags -eq 'ObjectAceTypePresent' )
                        {
                            $foundExtended = $true
                            $aceType       = $aceEntry.ObjectAceType
                            $extended      = Get-ExtendedAttribute $aceType
                            $AceString    += " [Extended Attribute = $extended]"
                        }
                    }

                    if( -not $foundExtended )
                    {
                        if($aceEntry.AccessMask)
                        {
                            $foundAccess = Get-AccessRights $aceEntry.AccessMask $Type
    
                            if($foundAccess)
                            {
                                $AceString += " ({0})" -f ($foundAccess -join ", ")
                            }
                        }    
                    }

                    $AceString
                }
            }
        }
    }

    Process
    {
        Set-StrictMode -Version Latest
 
        ## $rawSecurityDescriptor = [Security.AccessControl.CommonSecurityDescriptor]::new($false,$false,$Sddl)
        $rawSecurityDescriptor = New-Object -TypeName Security.AccessControl.CommonSecurityDescriptor `
                                    -ArgumentList $false, $false, $Sddl

        $owner   = ConvertTo-NtAccount $rawSecurityDescriptor.Owner
        $group   = ConvertTo-NtAccount $rawSecurityDescriptor.Group
        $control = ConvertTo-Control   $rawSecurityDescriptor.ControlFlags.value__
        $dAcl    = ConvertTo-AceString $rawSecurityDescriptor.DiscretionaryAcl $Type
        $sAcl    = ConvertTo-AceString $rawSecurityDescriptor.SystemAcl $Type

        [PSCustomObject] @{
            Owner            = $owner
            Group            = $group
            ControlFlags     = $control
            DiscretionaryAcl = @( $dAcl )
            SystemAcl        = @( $sAcl )
            RawDescriptor    = $rawSecurityDescriptor
        }
    }
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region for checking elevation
Function ElevatedSession
{
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

	If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
	{
		Write-Verbose "$(Get-Date): This is an elevated PowerShell session"
		Return $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): This is NOT an elevated PowerShell session"
		Return $False
	}
}

Function CheckElevation
{
	Write-Verbose "$(Get-Date): Testing for elevated PowerShell session."
	#see if session is elevated
	$Elevated = ElevatedSession
	
	If($Elevated -eq $False)
	{
		#abort script
		Write-Error "
		`n`n
		`t`t
		The Citrix FAS PowerShell cmdlets require an elevated PowerShell session.
		`n`n
		`t`t
		Rerun the script from an elevated PowerShell session. The script will now close.
		`n`n
		"
		Write-Verbose "$(Get-Date): "
		AbortScript
	}
}
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	If(!(Check-NeededPSSnapins "Citrix.Authentication.FederatedAuthenticationService.V1"))

	{
		#We're missing Citrix Snapins that we need
		Write-Error "
		`n`n
		`t`t
		Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. 
		`n`n
		`t`t
		Are you sure you are running this script against a Citrix FAS Server? 
		`n`n
		`t`t
		If you are running the script remotely, did you install Studio or the PowerShell snapins on $($env:computername)?
		`n`n
		`t`t
		Please see the Prerequisites section in the ReadMe file (https://carlwebster.sharefile.com/d-s8e92231489542428).
		`n`n
		`t`t
		Script will now close.
		"
		AbortScript
	}
	
	[string]$Script:FASDisplayName = Get-RegistryValue2 "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Citrix User Credential Service" "DisplayName" $env:ComputerName

	If($Null -eq $Script:FASDisplayName)
	{
		$Script:FASDisplayName = "Unable to determine"
	}

	$Script:Title = "Citrix FAS Inventory"
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdpath\FASV1InventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime         : $($AddDateTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Citrix Templates Only: $($AddDateTime)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name         : $($Script:CoName)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address      : $($CompanyAddress)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email        : $($CompanyEmail)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax          : $($CompanyFax)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone        : $($CompanyPhone)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page           : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev                  : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile         : $($Script:DevErrorFile)" 4>$Null
		}
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word FileName        : $($Script:WordFileName)" 4>$Null
		}
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTML FileName        : $($Script:HtmlFileName)" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDF Filename         : $($Script:PDFFileName)" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Text FileName        : $($Script:TextFileName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder               : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From                 : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Hardware Inventory   : $($Hardware)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Limit User Certs     : $($LimitUserCertificates)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                  : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML         : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF          : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT         : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD         : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info          : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port            : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server          : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title                : $($Script:FASDisplayName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                   : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL              : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name            : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected          : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version         : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture            : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture          : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language        : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version         : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start         : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time         : $($Str)" 4>$Null
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}
	#cleanup obj variables
	$Script:Output = $Null
}
#endregion

#region process root certificate
Function ProcessRootCA
{
	Write-Verbose "$(Get-Date): Retrieving Root Certificate Information"
	$CA = (Get-FasMsCertificateAuthority -EA 0 | Where-Object {$_.IsDefault -eq $True}).Address
	
	If(!($?) -or ($Null -eq $CA))
	{
		Write-Error "
		`n`n
		`t`t
		Unable to retrieve the Certificate Authority data from FAS.
		`n`n
		`t`t
		The script will now close.
		`n`n
		"
		Write-Verbose "$(Get-Date): "
		
		If(($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
		{
			$AbstractTitle = "Citrix FAS Inventory"
			$SubjectTitle = "Citrix FAS Inventory"
			UpdateDocumentProperties $AbstractTitle $SubjectTitle
		}
		ProcessDocumentOutput "Abort"
		AbortScript
	}
	
	Write-Verbose "$(Get-Date): Retrieving data for $($CA) in cert:\CurrentUser\CA"
	$TmpArray = $CA.Split("\")
	$CAServer = $TmpARray[0]
	$CAName = $TmpArray[1]
	$Results = Invoke-Command -ScriptBlock {Get-ChildItem -path cert:\CurrentUser\CA} -computername $CAServer | Where-Object {$_.Issuer -Like "*$CAName*"} | Select-Object Issuer,NotBefore,NotAfter,Subject

	If(!($?))
	{
		Write-Warning "$(Get-Date): Unable to retrieve Root Certificate information"
		$IssuedTo = "Unable to retrieve Root Certificate information"
		$IssuedBy = "Unable to retrieve Root Certificate information"
		$ValidNotBefore = (Get-Date).ToShortDateString()
		$ValidNotAfter = (Get-Date).ToShortDateString()
	}
	ElseIf($? -and $null -eq $Results)
	{
		Write-Warning "$(Get-Date): No Root Certificate found issued by $CAName"
		$IssuedTo = "No Root Certificate information found issued by $CAName"
		$IssuedBy = "No Root Certificate information found issued by $CAName"
		$ValidNotBefore = (Get-Date).ToShortDateString()
		$ValidNotAfter = (Get-Date).ToShortDateString()
	}
	Else
	{
		If($Results -is [array])
		{
			$tmp = $Results[0].Subject
			$start = $tmp.IndexOf("=")
			$Stop = $tmp.IndexOf(",")
			$IssuedTo = $tmp.SubString($Start+1,(($Stop)-($Start+1)))

			$tmp = $Results[0].Issuer
			$start = $tmp.IndexOf("=")
			$Stop = $tmp.IndexOf(",")
			$IssuedBy = $tmp.SubString($Start+1,(($Stop)-($Start+1)))

			$ValidNotBefore = $Results[0].NotBefore.ToShortDateString()
			$ValidNotAfter = $Results[0].NotAfter.ToShortDateString()
		}
		Else
		{
			$tmp = $Results.Subject
			$start = $tmp.IndexOf("=")
			$Stop = $tmp.IndexOf(",")
			$IssuedTo = $tmp.SubString($Start+1,(($Stop)-($Start+1)))

			$tmp = $Results.Issuer
			$start = $tmp.IndexOf("=")
			$Stop = $tmp.IndexOf(",")
			$IssuedBy = $tmp.SubString($Start+1,(($Stop)-($Start+1)))

			$ValidNotBefore = $Results.NotBefore.ToShortDateString()
			$ValidNotAfter = $Results.NotAfter.ToShortDateString()
		}
	}
	
	$CAobj = [PSCustomObject] @{
		CA                 = $CA
		CAServer           = $CAServer
		CAName             = $CAName
		IssuedTo           = $IssuedTo
		IssuedBy           = $IssuedBy
		ValidFrom          = $ValidNotBefore
		ValidTo            = $ValidNotAfter
	}
	
	OutputRootCA $CAObj
	
	$CAObj = $Null
}

Function OutputRootCA
{
	Param([object] $RCAObj)
	
	Write-Verbose "$(Get-Date): Add Root Certificate Information"
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "Root Certificate Information"
		$ScriptInformation = New-Object System.Collections.ArrayList
	}
	If($Text)
	{
		Line 0 "Root Certificate Information"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Root Certificate Information"
		$rowdata = @()
	}

	ForEach($Obj1 in $RCAObj)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation.Add(@{Data = "Certificate authority server"; Value = $Obj1.CAServer; }) > $Null
			$ScriptInformation.Add(@{Data = "Certificate authority name"; Value = $Obj1.CAName; }) > $Null
			$ScriptInformation.Add(@{Data = 'Issued to'; Value = $Obj1.IssuedTo; }) > $Null
			$ScriptInformation.Add(@{Data = 'Issued by'; Value = $Obj1.IssuedBy; }) > $Null
			$ScriptInformation.Add(@{Data = 'Valid from'; Value = "$($Obj1.ValidFrom) to $($Obj1.ValidTo)"; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Certificate authority server`t: " $Obj1.CAServer
			Line 1 "Certificate authority name`t: " $Obj1.CAName
			Line 1 "Issued to`t`t`t: " $Obj1.IssuedTo
			Line 1 "Issued by`t`t`t: " $Obj1.IssuedBy
			Line 1 "Valid from $($Obj1.ValidFrom) to $($Obj1.ValidTo)"
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("Certificate authority server",($Script:htmlsb),$Obj1.CAServer,$htmlwhite)
			$rowdata += @(,('Certificate authority name',($Script:htmlsb),$Obj1.CAName,$htmlwhite))
			$rowdata += @(,('Issued to',($Script:htmlsb),$Obj1.IssuedTo,$htmlwhite))
			$rowdata += @(,('Issued by',($Script:htmlsb),$Obj1.IssuedBy,$htmlwhite))
			$rowdata += @(,('Valid from',($Script:htmlsb),"$($Obj1.ValidFrom) to $($Obj1.ValidTo)",$htmlwhite))

			$msg = ""
			$columnWidths = @("150","150")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
		
		If($Hardware)
		{
			GetComputerWMIInfo $Obj1.CAServer
		}
	}
}
#endregion

#region process certificate authorities
Function ProcessCAs
{
	Write-Verbose "$(Get-Date): Retrieving Certificate Authorities Information"
	$FASCAs = New-Object System.Collections.ArrayList

	$CAs = Get-FasMsCertificateAuthority -EA 0
	
	If(!$?)
	{
		Write-Warning "$(Get-Date): Error - Unable to retrieve Certificate Authorities information"
		$FASCAs = "Error - Unable to retrieve Certificate Authorities information"
	}
	ElseIf($? -and $null -eq $CAs)
	{
		Write-Warning "$(Get-Date): Warning - No Certificate Authorities found"
		$FASCAs = "Warning - No Certificate Authorities found"
	}
	ElseIf($? -and $null -ne $CAs)
	{
		ForEach($CA in $CAs)
		{
			If($CitrixTemplatesOnly)
			{
				$PTs = $CA.PublishedTemplates | Where-Object {$_ -like "*citrix*"}
			}
			Else
			{
				$PTs = $CA.PublishedTemplates
			}

			$PTs = $PTs | Sort-Object
			
			$PTACLInfo = New-Object System.Collections.ArrayList

			ForEach($PT in $PTs)
			{
				$PTInfo = Get-FasMsTemplate -Name $PT -EA 0
				
				$PT_SDDL = ConvertFrom-SDDL-MBS $PTInfo.ACL
				$PT_SDDL.DiscretionaryACL = $PT_SDDL.DiscretionaryACL | Sort-Object
				
				$PTObj = [PSCustomObject] @{
					PTName            = $PT
					PTACLOwner        = $PT_SDDL.Owner
					PTACLGroup        = $PT_SDDL.Group
					PTACLControlFlags = $PT_SDDL.ControlFlags
					PTACL             = $PT_SDDL.DiscretionaryACL
				}
				$null = $PTACLInfo.Add($PTObj)
			}
			
			$CAobj = [PSCustomObject] @{
				CA                   = $CA.Address
				CAIsAccessible       = $CA.IsAccessible
				CAIsDefault          = $CA.IsDefault
				CAPublishedTemplates = $PTs
				CAPTInfo             = $PTACLInfo
			}
			$null = $FASCAs.Add($CAObj)
		}
	}
	
	OutputCAs $FASCAs
	
	$FASCAs    = $Null
	$PTACLInfo = $Null
	$CAObj     = $Null
}

Function OutputCAs
{
	Param([array] $FASCAs)
	Write-Verbose "$(Get-Date): Output Certificate Authorities Information"
	
	If($FASCAs -like "*error*" -or $FASCAs -like "*warning*")
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "Certificate Authorities Information"
			WriteWordline 0 0 "$($FASCAs)"
		}
		If($Text)
		{
			Line 0 "Certificate Authorities Information"
			line 0 "$($FASCAs)"
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "Certificate Authoritiess Information"
			WriteHTMLline 0 0 "$($FASCAs)"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "Certificate Authorities Information"
		}
		If($Text)
		{
			Line 0 "Certificate Authorities Information"
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "Certificate Authorities Information"
		}

		$First = $True
		ForEach($item in $FASCAs)
		{
			If($MSWord -or $PDF)
			{
				If(!$First)
				{
					$Script:Selection.InsertNewPage()
				}
				Else
				{
					$First = $False
				}
				WriteWordLine 2 0 "Certificate Authority " $item.CA
			}
			If($Text)
			{
				Line 0 "Certificate Authority " $item.CA
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "Certificate Authority " $item.CA
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Address"; Value = $item.CA; }) > $Null
				$ScriptInformation.Add(@{Data = "Is accessible"; Value = $item.CAIsAccessible; }) > $Null
				$ScriptInformation.Add(@{Data = "Is default"; Value = $item.CAIsDefault; }) > $Null
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 75;
				$Table.Columns.Item(2).Width = 200;

				SetWordCellFormat -Collection $Table -Size 9
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				WriteWordLine 2 0 "Published Templates"
				ForEach($PT in $item.CAPTInfo)
				{
					$ScriptInformation = New-Object System.Collections.ArrayList
					$ScriptInformation.Add(@{Data = "Template name"; Value = $PT.PTName; }) > $Null
					$ScriptInformation.Add(@{Data = "ACL Owner"; Value = $PT.PTACLOwner; }) > $Null
					$ScriptInformation.Add(@{Data = "ACL Group"; Value = $PT.PTACLGroup; }) > $Null
					$ScriptInformation.Add(@{Data = "ACL Control Flags"; Value = $PT.PTACLControlFlags; }) > $Null
					$cnt = -1
					ForEach($xACL in $PT.PTACL)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							$ScriptInformation.Add(@{Data = "ACL"; Value = $xACL; }) > $Null
						}
						Else
						{
							$ScriptInformation.Add(@{Data = ""; Value = $xACL; }) > $Null
						}
					}
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 80;
					$Table.Columns.Item(2).Width = 420;

					SetWordCellFormat -Collection $Table -Size 9
					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}

			}
			If($Text)
			{
				Line 1 "Address`t`t`t: " $item.CA
				Line 1 "Is accessible`t`t: " $item.CAIsAccessible
				Line 1 "Is default`t`t: " $item.CAIsDefault
				Line 0 ""

				Line 1 "Published Templates"
				ForEach($PT in $item.CAPTInfo)
				{
					Line 2 "Template name: " $PT.PTName
					Line 3 "ACL Owner`t : " $PT.PTACLOwner
					Line 3 "ACL Group`t : " $PT.PTACLGroup
					Line 3 "ACL Control Flags: " $PT.PTACLControlFlags
					$cnt = -1
					ForEach($xACL in $PT.PTACL)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							Line 3 "ACL`t`t : " $xACL
						}
						Else
						{
							Line 5 "   " $xACL
						}
					}
					Line 0 ""
				}
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Address",($Script:htmlsb),$item.CA,$htmlwhite)
				$rowdata += @(,('Is accessible',($Script:htmlsb),$item.CAIsAccessible,$htmlwhite))
				$rowdata += @(,('Is default',($Script:htmlsb),$item.CAIsDefault,$htmlwhite))
				
				$msg = ""
				$columnWidths = @("100","250")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths 
				WriteHTMLLine 0 0 ""

				WriteHTMLLine 2 0 "Published Templates"
				ForEach($PT in $item.CAPTInfo)
				{
					$rowdata = @()
					$columnHeaders = @("Template name",($Script:htmlsb),$PT.PTName,$htmlwhite)
					$rowdata += @(,("ACL Owner",($Script:htmlsb),$PT.PTACLOwner,$htmlwhite))
					$rowdata += @(,("ACL Group",($Script:htmlsb),$PT.PTACLGroup,$htmlwhite))
					$rowdata += @(,("ACL Control Flags",($Script:htmlsb),$PT.PTACLControlFlags,$htmlwhite))
					$cnt = -1
					ForEach($xACL in $PT.PTACL)
					{
						$cnt++
						
						If($cnt -eq 0)
						{
							$rowdata += @(,("ACL",($Script:htmlsb),$xACL,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,("",($Script:htmlsb),$xACL,$htmlwhite))
						}
					}
					$msg = ""
					$columnWidths = @("100","500")
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths 
					WriteHTMLLine 0 0 ""
				}
			}
		}
	}
}

#endregion

#region process FAS server
Function ProcessFASServer
{
	Write-Verbose "$(Get-Date): Retrieving FAS server Information"
	$results = Get-FasServer -EA 0
	
	If(!$? -or $null -eq $results)
	{
		Write-Error "
		`n`n
		`t`t
		Unable to retrieve the FAS server.
		`n`n
		`t`t
		The script cannot run without FAS server data. The script will now close.
		`n`n
		"
		Write-Verbose "$(Get-Date): "
		If(($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
		{
			$AbstractTitle = "Citrix FAS Inventory"
			$SubjectTitle = "Citrix FAS Inventory"
			UpdateDocumentProperties $AbstractTitle $SubjectTitle
		}
		ProcessDocumentOutput "Abort"
		AbortScript
	}
	
	$Script:CitrixFasAddress = $results[0].Address
	$FASServers = New-Object System.Collections.ArrayList
	
	ForEach($result in $results)
	{
		#get the installed product version n each FAS server
		$FASVersion = Get-RegistryValue2 "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Citrix User Credential Service" "DisplayName" $result.Address
		
	
		#get the data for AdministrationACL
		$SDDL = ConvertFrom-SDDL-MBS $result.AdministrationACL
		
		$FASobj = [PSCustomObject] @{
			Address             = $result.Address
			Index               = $result.Index
			Version             = $result.Version
			FASVersion          = $FASVersion
			MaintenanceMode     = $result.MaintenanceMode
			AdministrationACL   = $result.AdministrationACL
			Capabilities        = $result.Capabilities
			ACLOwner            = $SDDL.Owner
			ACLGroup            = $SDDL.Group
			ACLControlFlags     = $SDDL.ControlFlags
			ACLDiscretionaryACL = $SDDL.DiscretionaryACL
		}
		$null = $FASServers.Add($FASobj)
	}
	
	OutputFASServer $FASServers
	
	$FASServers = $Null
	$FASObj     = $Null
}

Function OutputFASServer
{
	Param([array]$FASServers)
	
	Write-Verbose "$(Get-Date): Output FAS server Information"
	$FASServers = $FASServers | Sort-Object Address
	
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
		WriteWordLine 1 0 "FAS Server Information"
		$ScriptInformation = New-Object System.Collections.ArrayList
	}
	If($Text)
	{
		Line 0 "FAS Server Information"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "FAS Server Information"
		$rowdata = @()
	}

	$First = $True
	ForEach($item in $FASServers)
	{
		If($MSWord -or $PDF)
		{
			If(!$First)
			{
				$Script:Selection.InsertNewPage()
			}
			Else
			{
				$First = $False
			}
			WriteWordLine 2 0 "FAS Server " $item.Address
		}
		If($Text)
		{
			Line 0 "FAS Server " $item.Address
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "FAS Server " $item.Address
		}
		
		$Capabilities = $item.Capabilities.Split(",")
		
		If($MSWord -or $PDF)
		{
			$ScriptInformation.Add(@{Data = "FAS Address"; Value = $item.Address; }) > $Null
			$ScriptInformation.Add(@{Data = 'Index'; Value = $item.Index; }) > $Null
			$ScriptInformation.Add(@{Data = 'Version'; Value = $item.Version; }) > $Null
			$ScriptInformation.Add(@{Data = 'FAS installed version'; Value = $item.FASVersion; }) > $Null
			$ScriptInformation.Add(@{Data = 'Maintenance mode'; Value = $item.MaintenanceMode; }) > $Null
			$ScriptInformation.Add(@{Data = 'Administration ACL'; Value = ""; }) > $Null
			$ScriptInformation.Add(@{Data = '     ACL Owner'; Value = $item.ACLOwner; }) > $Null
			$ScriptInformation.Add(@{Data = '     ACL Group'; Value = $item.ACLGroup; }) > $Null
			$ScriptInformation.Add(@{Data = '     ACL Control Flags'; Value = $item.ACLControlFlags; }) > $Null
			$ScriptInformation.Add(@{Data = '     Discretionary ACL'; Value = $item.ACLDiscretionaryACL; }) > $Null
			
			$cnt = -1
			ForEach($Capability in $Capabilities)
			{
				$cnt++
				
				If($cnt -eq 0)
				{
					$ScriptInformation.Add(@{Data = 'Capabilities'; Value = $Capability; }) > $Null
				}
				Else
				{
					$ScriptInformation.Add(@{Data = ''; Value = $Capability; }) > $Null
				}
			}

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 350;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "FAS Address`t`t: " $item.Address
			Line 1 "Index`t`t`t: " $item.Index
			Line 1 "Version`t`t`t: " $item.Version
			Line 1 "FAS installed version`t: " $item.FASVersion
			Line 1 "Maintenance mode`t: " $item.MaintenanceMode
			Line 1 "Administration ACL`t: " 
			Line 2 "ACL Owner`t`t: " $item.ACLOwner
			Line 2 "ACL Group`t`t: " $item.ACLGroup
			Line 2 "ACL Control Flags`t: " $item.ACLControlFlags
			Line 2 "Discretionary ACL`t: " $item.ACLDiscretionaryACL

			$cnt = -1
			ForEach($Capability in $Capabilities)
			{
				$cnt++
				
				If($cnt -eq 0)
				{
					Line 1 "Capabilities`t`t: " $Capability
				}
				Else
				{
					Line 4 ": " $Capability
				}
			}
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("FAS Address",($Script:htmlsb),$item.Address,$htmlwhite)
			$rowdata += @(,('Index',($Script:htmlsb),$item.Index.ToString(),$htmlwhite))
			$rowdata += @(,('Version',($Script:htmlsb),$item.Version,$htmlwhite))
			$rowdata += @(,('FAS installed version',($Script:htmlsb),$item.FASVersion,$htmlwhite))
			$rowdata += @(,('Maintenance mode',($Script:htmlsb),$item.MaintenanceMode.ToString(),$htmlwhite))
			$rowdata += @(,('Administration ACL',($Script:htmlsb),"",$htmlwhite))
			$rowdata += @(,('     ACL Owner',($Script:htmlsb),$item.ACLOwner,$htmlwhite))
			$rowdata += @(,('     ACL Group',($Script:htmlsb),$item.ACLGroup,$htmlwhite))
			$rowdata += @(,('     ACL Control Flags',($Script:htmlsb),$item.ACLControlFlags,$htmlwhite))
			$rowdata += @(,('     Discretionary ACL',($Script:htmlsb),$item.ACLDiscretionaryACL,$htmlwhite))
			
			$cnt = -1
			ForEach($Capability in $Capabilities)
			{
				$cnt++
				
				If($cnt -eq 0)
				{
					$rowdata += @(,('Capabilities',($Script:htmlsb),$Capability,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($Script:htmlsb),$Capability,$htmlwhite))
				}
			}

			$msg = ""
			$columnWidths = @("150","400")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
		
		If($Hardware)
		{
			GetComputerWMIInfo $item.Address
		}
	}
}
#endregion

#region process local admin policy
Function ProcessLocalAdminPolicy
{
	$results = Get-FasAdministrationPolicy -EA 0
	
	If(!($?))
	{
		#error
		Write-Warning "$(Get-Date): Error retrieving FAS Local Admin Policy"
	}
	ElseIf($? -and $Null -eq $results)
	{
		#nothing returned
		Write-Warning "$(Get-Date): Warning. No FAS Local Admin Policy was found"
	}
	ElseIf($? -and $Null -ne $results)
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "FAS Local Administration Policy "
			WriteWordLine 2 0 "FAS Local Administration Policy "
			$ScriptInformation = New-Object System.Collections.ArrayList
		}
		If($Text)
		{
			Line 0 "FAS Local Administration Policy "
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "FAS Local Administration Policy "
			WriteHTMLLine 2 0 "FAS Local Administration Policy "
			$rowdata = @()
		}

		If($MSWord -or $PDF)
		{
			$ScriptInformation.Add(@{Data = "Default to local host"; Value = $results.DefaultToLocalhost; }) > $Null
			$ScriptInformation.Add(@{Data = 'Check address against GPO'; Value = $results.CheckAddressAgainstGpo; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Default to local host`t : " $results.DefaultToLocalhost
			Line 1 "Check address against GPO: " $results.CheckAddressAgainstGpo
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("Default to local host",($Script:htmlsb),$results.DefaultToLocalhost.ToString(),$htmlwhite)
			$rowdata += @(,('Check address against GPO',($Script:htmlsb),$results.CheckAddressAgainstGpo.ToString(),$htmlwhite))

			$msg = ""
			$columnWidths = @("150","150")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region process Private Key Pool Info
Function ProcessPrivateKeyPoolInfo
{
	$results = Get-FasPrivateKeyPoolInfo -EA 0
	
	If(!($?))
	{
		#error
		Write-Warning "$(Get-Date): Error retrieving FAS Private Key Pool Info"
	}
	ElseIf($? -and $Null -eq $results)
	{
		#nothing returned
		Write-Warning "$(Get-Date): Warning. No FAS Private Key Pool Info was found"
	}
	ElseIf($? -and $Null -ne $results)
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "FAS Private Key Pool Info "
			WriteWordLine 2 0 "FAS Private Key Pool Info "
			$ScriptInformation = New-Object System.Collections.ArrayList
		}
		If($Text)
		{
			Line 0 "FAS Private Key Pool Info "
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "FAS Private Key Pool Info "
			WriteHTMLLine 2 0 "FAS Private Key Pool Info "
			$rowdata = @()
		}

		If($MSWord -or $PDF)
		{
			$ScriptInformation.Add(@{Data = "Target size"; Value = $results.TargetSize; }) > $Null
			$ScriptInformation.Add(@{Data = 'Current size'; Value = $results.CurrentSize; }) > $Null

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 150;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Target size : " $results.TargetSize
			Line 1 "Current size: " $results.CurrentSize
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("Target size",($Script:htmlsb),$results.TargetSize.ToString(),$htmlwhite)
			$rowdata += @(,('Current size',($Script:htmlsb),$results.CurrentSize.ToString(),$htmlwhite))

			$msg = ""
			$columnWidths = @("150","150")
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region process FAS rules
Function ProcessFASRules
{
	Write-Verbose "$(Get-Date): Retrieving FAS rules Information"
	$results = Get-FasRule -EA 0
	
	If(!($?))
	{
		#error
		$FASRules = "Error retrieving FAS rules Information"
		Write-Warning "$(Get-Date): Error retrieving FAS rules Information"
	}
	ElseIf($? -and $Null -eq $results)
	{
		#nothing returned
		$FASRules = "Warning. No FAS rules were found"
		Write-Warning "$(Get-Date): Warning. No FAS rules were found"
	}
	ElseIf($? -and $Null -ne $results)
	{
		$FASRules = New-Object System.Collections.ArrayList
		
		ForEach($result in $results)
		{
		
			#get the data for StoreFrontACL
			$SF_SDDL = ConvertFrom-SDDL-MBS $result.StoreFrontAcl
			
			#get the data for VdaACL
			$VDA_SDDL = ConvertFrom-SDDL-MBS $result.VdaAcl
			
			#get the data for StoreFrontACL
			$User_SDDL = ConvertFrom-SDDL-MBS $result.UserAcl
			
			#get the data for AdministrationACL
			$Admin_SDDL = ConvertFrom-SDDL-MBS $result.AdministrationACL
			
			$FASRuleDefinitions = New-Object System.Collections.ArrayList
			$FASRuleDefinitionCAs = New-Object System.Collections.ArrayList

			ForEach($Definition in $result.CertificateDefinitions)
			{
				#get certificate definition
				$CertDefinition = Get-FasCertificateDefinition -Name $Definition
				
				#get CAs
				ForEach($CA in $CertDefinition.CertificateAuthorities)
				{
					$null = $FASRuleDefinitionCAs.Add($CA)
				}

				$CertDefObj = [PSCustomObject] @{
					CertDefCA = $FASRuleDefinitionCAs
					CertDefMsTemplate = $CertDefinition.MsTemplate
					CertDefAvailableAfterLogon = $CertDefinition.InSession
					}

				$null = $FASRuleDefinitions.Add($CertDefObj)
			}
		
			$Rulesobj = [PSCustomObject] @{
				RuleName                     = $result.Name
				RuleCertificateDefinitions   = $result.CertificateDefinitions
				
				RuleDefinitions              = $FASRuleDefinitions
				
				RuleSFAcl                    = $result.StoreFrontAcl
				RuleSFAclOwner               = $SF_SDDL.Owner
				RuleSFAclGroup               = $SF_SDDL.Group
				RuleSFAclControlFlags        = $SF_SDDL.ControlFlags
				RuleSFAclDiscretionaryACL    = $SF_SDDL.DiscretionaryACL
				
				RuleVDAAcl                   = $result.VdaAcl
				RuleVDAAclOwner              = $VDA_SDDL.Owner
				RuleVDAAclGroup              = $VDA_SDDL.Group
				RuleVDAAclControlFlags       = $VDA_SDDL.ControlFlags
				RuleVDAAclDiscretionaryACL   = $VDA_SDDL.DiscretionaryACL
				
				RuleUserAcl                  = $result.UserAcl
				RuleUserAclOwner             = $User_SDDL.Owner
				RuleUserAclGroup             = $User_SDDL.Group
				RuleUserAclControlFlags      = $User_SDDL.ControlFlags
				RuleUserAclDiscretionaryACL  = $User_SDDL.DiscretionaryACL
				
				RuleAdministrationACL        = $result.AdministrationACL
				RuleAdminACLOwner            = $Admin_SDDL.Owner
				RuleAdminACLGroup            = $Admin_SDDL.Group
				RuleAdminAclControlFlags     = $Admin_SDDL.ControlFlags
				RuleAdminACLDiscretionaryACL = $Admin_SDDL.DiscretionaryACL
			}
			$null = $FASRules.Add($Rulesobj)
		}
	}
	OutputFASRules $FASRules
	
	$FASRules             = $Null
	$FASRuleDefinitions   = $Null
	$FASRuleDefinitionCAs = $Null
	$CertDefObj           = $Null
	$Rulesobj             = $Null
}

Function OutputFASRules
{
	Param([array]$FASRules)
	
	Write-Verbose "$(Get-Date): Output FAS Rules Information"
	
	If($FASRules -like "*error*" -or $FASRules -like "*warning*")
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "FAS Rules Information"
			WriteWordline 0 0 "$($FASRules)"
		}
		If($Text)
		{
			Line 0 "FAS Rules Information"
			line 0 "$($FASRules)"
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "FAS Rules Information"
			WriteHTMLline 0 0 "$($FASRules)"
		}
	}
	Else
	{
		$FASRules = $FASRules | Sort-Object RuleName
		
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "FAS Rules Information"
		}
		If($Text)
		{
			Line 0 "FAS Rules Information"
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "FAS Rules Information"
		}

		$First = $True
		ForEach($item in $FASRules)
		{
			#unpack rule definitions and rule CAs
			If($MSWord -or $PDF)
			{
				If(!$First)
				{
					$Script:Selection.InsertNewPage()
				}
				Else
				{
					$First = $False
				}
				WriteWordLine 2 0 "FAS Rule " $item.RuleName
			}
			If($Text)
			{
				Line 0 "FAS Rule " $item.RuleName
			}
			If($HTML)
			{
				WriteHTMLLine 2 0 "FAS Rule " $item.RuleName
			}
			
			If($MSWord -or $PDF)
			{
				$ScriptInformation = New-Object System.Collections.ArrayList
				$ScriptInformation.Add(@{Data = "Rule name"; Value = $item.RuleName; }) > $Null
				$ScriptInformation.Add(@{Data = "Certificate Authority"; Value = $item.RuleDefinitions.CertDefCA; }) > $Null
				$cnt = -1
				ForEach($tmp in $item.RuleDefinitions.CertDefCA)
				{
					$cnt++
					
					If($cnt -ge 1)
					{
						$ScriptInformation.Add(@{Data = ""; Value = $tmp; }) > $Null
					}
				}
				$ScriptInformation.Add(@{Data = "Certificate Template"; Value = $item.RuleDefinitions.CertDefMsTemplate; }) > $Null
				$ScriptInformation.Add(@{Data = "Available after logon"; Value = $item.RuleDefinitions.CertDefAvailableAfterLogon; }) > $Null
				$ScriptInformation.Add(@{Data = "Security Access Control Lists"; Value = ""; }) > $Null
				
				$cnt = -1
				ForEach($SFServer in $item.RuleSFAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "     List of StoreFront servers that can use this rule"; Value = $SFServer; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $SFServer; }) > $Null
					}
				}
				
				$cnt = -1
				ForEach($VDA in $item.RuleVDAAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "     List of VDAs that can be logged into by this rule"; Value = $VDA; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $VDA; }) > $Null
					}
				}
				
				$cnt = -1
				ForEach($User in $item.RuleUserAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$ScriptInformation.Add(@{Data = "     List of users that StoreFront can log in using this rule"; Value = $User; }) > $Null
					}
					Else
					{
						$ScriptInformation.Add(@{Data = ""; Value = $User; }) > $Null
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 225;
				$Table.Columns.Item(2).Width = 275;

				SetWordCellFormat -Collection $Table -Size 9
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Rule name`t`t: " $item.RuleName
				Line 1 "Certificate Authority`t: " $item.RuleDefinitions.CertDefCA
				$cnt = 0
				ForEach($tmp in $item.RuleDefinitions.CertDefCA)
				{
					$cnt++
					
					If($cnt -gt 1)
					{
						Line 4 "  " $tmp
					}
				}
				Line 1 "Certificate Template`t: " $item.RuleDefinitions.CertDefMsTemplate
				Line 1 "Available after logon`t: " $item.RuleDefinitions.CertDefAvailableAfterLogon
				Line 1 "Security Access Control Lists"
				
				$cnt = -1
				ForEach($SFServer in $item.RuleSFAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						Line 2 "List of StoreFront servers that can use this rule`t : " $SFServer
					}
					Else
					{
						Line 9 "   " $SFServer
					}
				}
				
				$cnt = -1
				ForEach($VDA in $item.RuleVDAAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						Line 2 "List of VDAs that can be logged into by this rule`t : " $VDA
					}
					Else
					{
						Line 9 "   " $VDA
					}
				}
				
				$cnt = -1
				ForEach($User in $item.RuleUserAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						Line 2 "List of users that StoreFront can log in using this rule : " $User
					}
					Else
					{
						Line 9 "   " $User
					}
				}
				
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Rule name",($Script:htmlsb),$item.RuleName,$htmlwhite)
				$rowdata += @(,('Certificate Authority',($Script:htmlsb),$item.RuleDefinitions.CertDefCA,$htmlwhite))
				$cnt = 0
				ForEach($tmp in $item.RuleDefinitions.CertDefCA)
				{
					$cnt++
					
					If($cnt -gt 1)
					{
						$rowdata += @(,('',($Script:htmlsb),$tmp,$htmlwhite))
					}
				}
				$rowdata += @(,('Certificate Template',($Script:htmlsb),$item.RuleDefinitions.CertDefMsTemplate,$htmlwhite))
				$rowdata += @(,('Available after logon',($Script:htmlsb),$item.RuleDefinitions.CertDefAvailableAfterLogon.ToString(),$htmlwhite))
				$rowdata += @(,('Security Access Control Lists',($Script:htmlsb),"",$htmlwhite))
				
				$cnt = -1
				ForEach($SFServer in $item.RuleSFAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$rowdata += @(,('     List of StoreFront servers that can use this rule',($Script:htmlsb),$SFServer,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',($Script:htmlsb),$SFServer,$htmlwhite))
					}
				}
				
				$cnt = -1
				ForEach($VDA in $item.RuleVDAAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$rowdata += @(,('     List of VDAs that can be logged into by this rule',($Script:htmlsb),$VDA,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',($Script:htmlsb),$VDA,$htmlwhite))
					}
				}
				
				$cnt = -1
				ForEach($User in $item.RuleUserAclDiscretionaryACL)
				{
					$cnt++
					
					If($cnt -eq 0)
					{
						$rowdata += @(,('     List of users that StoreFront can log in using this rule',($Script:htmlsb),$User,$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',($Script:htmlsb),$User,$htmlwhite))
					}
				}
				$msg = ""
				$columnWidths = @("325","400")
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths 
				WriteHTMLLine 0 0 ""
			}
		}
	}
}
#endregion

#region Process User Certificates
Function ProcessUserCertificates
{
	Write-Verbose "$(Get-Date): Retrieving User Certificate Information"
	$results = Get-FasUserCertificate -MaximumRecordCount $LimitUserCertificates -EA 0
	
	If(!($?))
	{
		#error
		$UserCerts = "Error retrieving User Certificate Information"
		Write-Warning "$(Get-Date): Error retrieving User Certificate Information"
	}
	ElseIf($? -and $Null -eq $results)
	{
		#nothing returned
		$UserCerts = "Warning. No User Certificates were found"
		Write-Warning "$(Get-Date): Warning. No User Certificates were found"
	}
	ElseIf($? -and $Null -ne $results)
	{
		$UserCerts = New-Object System.Collections.ArrayList

		ForEach($result in $results)
		{
			$UCertobj = [PSCustomObject] @{
				UPN            = $result.UserPrincipalName
				Role           = $result.Role
				CertDefinition = $result.CertificateDefinition
				ExpiryDate     = $result.ExpiryDate.ToString()
			}
			$null = $UserCerts.Add($UCertobj)
		}
	}
	OutputUserCertificates $UserCerts
	
	$UserCerts = $Null
	$UCertobj  = $Null
}

Function OutputUserCertificates
{
	Param([array]$UserCerts)
	
	Write-Verbose "$(Get-Date): Output User Certificate Information"
	
	If($UserCerts -like "*error*" -or $UserCerts -like "*warning*")
	{
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
			WriteWordLine 1 0 "User Certificate Information"
			WriteWordline 0 0 "$($UserCerts)"
		}
		If($Text)
		{
			Line 0 "User Certificate Information"
			line 0 "$($UserCerts)"
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "User Certificate Information"
			WriteHTMLline 0 0 "$($UserCerts)"
		}
	}
	Else
	{
		$UserCerts = $UserCerts | Sort-Object UPN
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 1 0 "User Certificate Information"
			$CertWordTable = @()
		}
		If($Text)
		{
			Line 0 "User Certificate Information"
			Line 1 "User Principal Name                                Role                      Certificate definition         Expiry Date           " 
			Line 1 "=================================================================================================================================="
			#       12345678901234567890123456789012345678901234567890S1234567890123456789012345S123456789012345678901234567890S1234567890123456789012
			#       50                                                 25                        30                             22
		}
		If($HTML)
		{
			WriteHTMLLine 1 0 "User Certificate Information"
			$rowdata = @()
		}
		
		ForEach($item in $UserCerts)
		{
			If($MSWord -or $PDF)
			{
				$CertWordTable += @{ 
				xUPN = $item.UPN;
				xRole = $item.Role;
				xCertDef = $item.CertDefinition;
				xExpiryDate = $item.ExpiryDate;
				}
			}
			If($Text)
			{
				Line 1 ( "{0,-50} {1,-25} {2,-30} {3,-22}" -f `
				$item.UPN, $item.Role, $item.CertDefinition, $item.ExpiryDate)
			}
			If($HTML)
			{
				$rowdata += @(,(
				$item.UPN,$htmlwhite,
				$item.Role,$htmlwhite,
				$item.CertDefinition,$htmlwhite,
				$item.ExpiryDate,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $CertWordTable `
			-Columns xUPN, xRole, xCertDef, xExpiryDate `
			-Headers "User Principal Name", "Role", "Certificate definition", "Expiry Date" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 75;
			$Table.Columns.Item(3).Width = 100;
			$Table.Columns.Item(3).Width = 125;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
			'User Principal Name',($Script:htmlsb),
			'Role',($Script:htmlsb),
			'Certificate definition',($Script:htmlsb),
			'Expiry Date',($Script:htmlsb))

			$msg = ""
			$columnWidths = @("250","100","100","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "650"

			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region script core
#Script begins

CheckElevation

ProcessScriptSetup

SetFilenames "CitrixFASInventory"

ProcessRootCA

ProcessCAs

ProcessFASServer

If(Get-Command Get-FasAdministrationPolicy -EA 0)
{
    ProcessLocalAdminPolicy
}

If(Get-Command Get-FasPrivateKeyPoolInfo -EA 0)
{
    ProcessPrivateKeyPoolInfo
}

ProcessFASRules

ProcessUserCertificates
#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

If(($MSWORD -or $PDF) -and ($Script:CoverPagesExist))
{
	$AbstractTitle = "$Script:FASDisplayName Inventory"
	$SubjectTitle = "Citrix FAS Inventory"
	UpdateDocumentProperties $AbstractTitle $SubjectTitle
}

ProcessDocumentOutput "Regular"

ProcessScriptEnd
#endregion