#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region Support

<#
.SYNOPSIS
    Creates a complete inventory of a Citrix NetScaler configuration using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix NetScaler configuration using Microsoft Word 
	and PowerShell.
	Creates a Word document named after the Citrix NetScaler Configuration.
	Document includes a Cover Page, Table of Contents, and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
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
		
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page if the Cover Page has the Address field.
	
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
	Company Email to use for the Cover Page if the Cover Page has the Email field.  
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page if the Cover Page has the Fax field.  
	
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
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022 at 6PM is 2022-06-01_1800.
	Output filename will be ReportName_2022-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	The default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER Log
	Generates a log file for troubleshooting.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry 
	Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry 
	Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry 
	Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry 
	Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\Docu-ADC.ps1 -CompanyName "Barry Schiffer 
	Consulting" -CoverPage "Mod" -UserName "Barry Schiffer"

	Will use:
		Barry Schiffer Consulting for the Company Name.
		Mod for the Cover Page format.
		Barry Schiffer for the User Name.
.EXAMPLE
	PS C:\PSScript .\Docu-ADC.ps1 -CN "Barry Schiffer Consulting" -CP 
	"Mod" -UN "Barry Schiffer"

	Will use:
		Barry Schiffer Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Barry Schiffer for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry 
	Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry 
	Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022 at 6PM is 2022-06-01_1800.
	Output filename will be Script_Template_2022-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Barry 
	Schiffer" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Barry 
	Schiffer"
	$env:username = Administrator

	Barry Schiffer for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2022 at 6PM is 2022-06-01_1800.
	Output filename will be Script_Template_2022-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript .\Docu-ADC.ps1 -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker 
	Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200"
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.
.EXAMPLE
	PS C:\PSScript .\Docu-ADC.ps1 -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Facet -UserName "Dr. Watson" -CompanyEmail 
	SuperSleuth@SherlockHolmes.com

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript >.\Docu-ADC.ps1 -Dev -ScriptInfo -Log
	
	Creates the default report.
	
	Creates a text file named NSInventoryScriptErrors_yyyyMMddTHHmmssffff.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named NSInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	NSDocScriptTranscript_yyyyMMddTHHmmssffff.txt.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -SmtpServer mail.domain.tld -From 
	XDAdmin@domain.tld -To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send an email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -SmtpServer smtp.office365.com -SmtpPort 
	587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADC.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, 
	the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word or PDF document.
.NOTES
	NAME: NetScaler_Script_v2_6_unsigned.ps1
	VERSION: 2.62
	AUTHOR: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters, Barry Schiffer
	LASTEDIT: February 18, 2022
#>

#endregion Support

#region script template
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "WordOrPDF") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False)] 
	[Alias("ADT")]
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From="",

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To="",

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False

	)
	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2014

<#
.NetScaler Documentation Script
    NAME: Docu-ADC.ps1
	VERSION NetScaler Script: 2.6
	AUTHOR NetScaler script: Barry Schiffer
    AUTHOR NetScaler script functions: Iain Brighton
    AUTHOR Script template: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
#Version 2.62 18-Feb-2022
#	Changed the date format for the transcript and error log files from yyyy-MM-dd_HHmm format to the FileDateTime format
#		The format is yyyyMMddTHHmmssffff (case-sensitive, using a 4-digit year, 2-digit month, 2-digit day, 
#		the letter T as a time separator, 2-digit hour, 2-digit minute, 2-digit second, and 4-digit millisecond). 
#		For example: 20221225T0840107271.
#	Fixed $Null comparisons that were on the wrong side
#	Fixed the German Table of Contents (Thanks to Rene Bigler)
#		From 
#			'de-'	{ 'Automatische Tabelle 2'; Break }
#		To
#			'de-'	{ 'Automatisches Verzeichnis 2'; Break }
#	In Function AbortScript, add test for the winword process and terminate it if it is running
#		Added stopping the transcript log if the log was enabled and started
#	In Functions AbortScript and SaveandCloseDocumentandShutdownWord, add code from Guy Leech to test for the "Id" property before using it
#	Replaced most script Exit calls with AbortScript to stop the transcript log if the log was enabled and started
#	Updated Functions CheckWordPrereq and SendEmail to the latest version
#	Updated the help text
#	Updated the ReadMe file
#
.Release Notes V2.61
#	Fix Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
#	Updated help text
.Release Notes V2.60
	Added -Dev and -ScriptInfo parameters
	Added four new Cover Page properties
		Company Address
		Company Email
		Company Fax
		Company Phone
	Added Function sendemail
	Added Log switch to create a transcript log
		Added function TranscriptLogging
	Fixed uninitialized variable for Admin Partitions
	Fixed French wording for Table of Contents 2 (Thanks to David Rouquier)
	Removed Hardware code as it is not needed for NetScaler
	Removed HTML and Text code as they are not used
	Removed code that made sure all Parameters were set to default values if for some reason they did exist, or values were $Null
	Removed ComputerName code as it is not needed for NetScaler
	Reordered the parameters in the help text and parameter list so they match and are grouped better
	Replaced _SetDocumentProperty function with Jim Moyle's Set-DocumentProperty function
	Updated Function ProcessScriptEnd for the new Cover Page properties, and Dev, ScriptInfo, Log Parameters
	Updated Function ShowScriptOptions for the new Cover Page properties, and Dev, ScriptInfo, Log Parameters
	Updated Function UpdateDocumentProperties for the new Cover Page properties and Parameters
	Updated Help Text
	Updated script to support Word 2016 but doing so removes support for Word 2007
.Release Notes version 2
    Overall
        Test group has grown from 5 to 20 people. A lot more testing on a lot more configs has been done.
        The result is that I've received a lot of nitty gritty bugs that are now solved. To many to list them all but this release is very very stable.
    New Script functionality
        New table function that now utilizes native word tables. Looks a lot better and is way faster
        Performance improvements; over 500% faster
        Better support for multi language Word versions. Will now always utilize cover page and TOC
    New NetScaler functionality:
        NetScaler Gateway
            Global Settings
            Virtual Servers settings and policies
            Policies Session/Traffic
	    NetScaler administration users and groups
        NetScaler Authentication
	        Policies LDAP / Radius
            Actions Local / RADIUS
            Action LDAP more configuration reported and changed table layout
        NetScaler Networking
            Channels
            ACL
        NetScaler Cache redirection
    Bugfixes
        Naming of items with spaces and quotes fixed
        Expressions with spaces, quotes, dashes and slashed fixed
        Grammatical corrections
        Rechecked all settings like enabled/disabled or on/off and corrected when necessary
        Time zone not show correctly when in GMT+....
        A lot more small items

.Release Notes version 1
    Version 1.0 supports the following NetScaler functionality:
	NetScaler System Information
	Version / NSIP / vLAN
	NetScaler Global Settings
	NetScaler Feature and mode state
	NetScaler Networking
	IP Address / vLAN / Routing Table / DNS
	NetScaler Authentication
	Local / LDAP
	NetScaler Traffic Domain
	Assigned Content Switch / Load Balancer / Service  / Server
	NetScaler Monitoring
	NetScaler Certificate
	NetScaler Content Switches
	Assigned Load Balancer / Service  / Server
	NetScaler Load Balancer
	Assigned Service  / Server
	NetScaler Service
	Assigned Server / monitor
	NetScaler Service Group
	Assigned Server / monitor
	NetScaler Server
	NetScaler Custom Monitor
	NetScaler Policy
	NetScaler Action
	NetScaler Profile

#>


Function AbortScript
{
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		If(Test-Path variable:global:word)
		{
			$Script:Word.quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()

	If($MSWord -or $PDF)
	{
		#is the winword Process still running? kill it

		#find out our session (usually "1" except on TS/RDC or Citrix)
		$SessionID = (Get-Process -PID $PID).SessionId

		#Find out if winword running in our session
		$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) | Select-Object -Property Id 
		If( $wordprocess -and $wordprocess.Id -gt 0)
		{
			Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
			Stop-Process $wordprocess.Id -EA 0
		}
	}
	
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $True) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Set-StrictMode -Version 2

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

#V2.60 added
If($Log) 
{
	#start transcript logging
	$Script:ThisScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$Script:LogPath = "$Script:ThisScriptPath\NSDocScriptTranscript_$(Get-Date -f FileDateTime).txt"
	
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
	$Script:DevErrorFile = "$($pwd.Path)\NSInventoryScriptErrors_$(Get-Date -f FileDateTime).txt"
}

If($Null -eq $MSWord)
{
	If($PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	AbortScript
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
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			AbortScript
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		AbortScript
	}
}

[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
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
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

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
	[int]$wdHeadingFormatFalse = 0 
	
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}
Else
{
	$Script:CoName = ""
}

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
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			#'de-'	{ 'Automatische Tabelle 2'; Break }
			'de-'	{ 'Automatisches Verzeichnis 2'; Break } #changed 18-feb-2022 rene bigler
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
#			'fr-'	{ 'Sommaire Automatique 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			# fix in 2.61 thanks to Johan Kallio 'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
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
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
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
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
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
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
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
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
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
		$ErrorActionPreference = $SaveEAPreference
		
		If(($MSWord -eq $False) -and ($PDF -eq $True))
		{
			Write-Host "`n`n`t`tThis script uses Microsoft Word's SaveAs PDF function, please install Microsoft Word`n`n"
			AbortScript
		}
		Else
		{
			Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
			AbortScript
		}
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = $null -ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
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

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

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
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
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

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Add DateTime    : $($AddDateTime)"
	Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
	Write-Verbose "$(Get-Date): Company Address : $($CompanyAddress)"
	Write-Verbose "$(Get-Date): Company Email   : $($CompanyEmail)"
	Write-Verbose "$(Get-Date): Company Fax     : $($CompanyFax)"
	Write-Verbose "$(Get-Date): Company Phone   : $($CompanyPhone)"
	Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
	Write-Verbose "$(Get-Date): Dev             : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Filename1       : $($Script:FileName1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $($Script:FileName2)"
	}
	Write-Verbose "$(Get-Date): Folder          : $($Folder)"
	Write-Verbose "$(Get-Date): From            : $($From)"
	Write-Verbose "$(Get-Date): Log             : $($Log)"
	Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
	Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
	Write-Verbose "$(Get-Date): To              : $($To)"
	Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
	Write-Verbose "$(Get-Date): User Name       : $($UserName)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $($RunningOS)"
	Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
	Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
	Write-Verbose "$(Get-Date): Word version    : $($Script:WordProduct)"
	Write-Verbose "$(Get-Date): Word language   : $($Script:WordLanguageValue)"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start    : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
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

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
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
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
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
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
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
						$CoverPage = "Línia lateral"
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
						$CoverPage = "Línea lateral"
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
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
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
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

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
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
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
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
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
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword Process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword running in our session
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) | Select-Object -Property Id 
	If( $wordprocess -and $wordprocess.Id -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord Process is still running. Attempting to stop WinWord Process # $($wordprocess.Id)"
		Stop-Process $wordprocess.Id -EA 0
	}
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $Script:Filename1 and $Script:Filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		If($DeliveryGroupsUtilization)
		{
			CheckExcelPreReq
		}

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
}

#Script begins

$script:startTime = Get-Date

#endregion script template

#region file name and title name
#The function SetFileName1andFileName2 needs your script output filename
SetFileName1andFileName2 "NetScaler Documentation"

#change title for your report
[string]$Script:Title = "NetScaler Documentation $CoName"
#endregion file name and title name

#region NetScaler Documentation Script Complete

$selection.InsertNewPage()

#region NetScaler Documentation Functions

<#
.SYNOPSIS
   Get a named property value from a string.
.DESCRIPTION
   Returns a case-insensitive property from a string, assuming the property is
   named before the actual property value and is separated by a space. For
   example, if the specified SearchString contained "-property1 <value1>
   -property2 <value2>�, searching for "-Property1" would return "<value1>".
.PARAMETER SearchString
   String to search for the specified property name.
.PARAMETER PropertyName
   The property name to search the SearchString for.
.PARAMETER Default
   If the property is not found returns the specified string. This parameter is
   optional and if not specified returns $null (by default) if the property is
   not found.
.PARAMETER RemoveQuotes
    Removes quotes from returned property values if present.
.PARAMETER ReplaceEscapedQuotes
    Replaces escaped quotes (\") with quotes (") from the returned property values
    if present. Note: This is generally used for display purposes only.
.EXAMPLE
   Get-StringProperty -SearchString $StringToSearch -PropertyName "-property1"

   This command searches the $StringToSearch variable for the presence of the property
   "-property1" and returns its value, if found. If the property name is not found,
   the default $null will be returned.
.EXAMPLE
   Get-StringProperty $StringToSearch "-property3" "Not found"

   This command searches the $StringToSearch variable for the presence of the property
   "-property3" and returns its value, if found. If the property name is not found,
   the "Not Found" string will be returned.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>


function Get-StringProperty {
    [CmdletBinding(HelpUri = 'http://virtualengine.co.uk/2014/searching-for-string-properties-with-powershell/')]
    [OutputType([String])]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString,
        # String property name to search for
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [Alias("Name","Property")] [string] $PropertyName,
        # Default return value for missing values
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=2)]
        [AllowNull()] [String] $Default = $null,
        # String delimiter, default to one or more spaces
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=3)]
        [ValidateNotNullOrEmpty()] [string] $Delimiter = ' ',
        # Remove quotes from quoted strings
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("NoQuotes")] [Switch] $RemoveQuotes,
        # Replace escaped quotes with quotes
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch] $ReplaceEscapedQuotes
    )

    Process {
        # First replace escaped quotes with '�'
        $SearchString = $SearchString.Replace('\"', "�");
        # Locate and replace quotes with '^^' and quoted spaces '^' to aid with parsing, until there are none left
        while ($SearchString.Contains('"')) {
            # Store the right-hand side temporarily, skipping the first quote
            $searchStringRight = $SearchString.Substring($SearchString.IndexOf('"') +1);
            # Extract the quoted text from the original string
            $quotedString = $SearchString.Substring($SearchString.IndexOf('"'), $searchStringRight.IndexOf('"') +2);
            # Replace the quoted text, replacing spaces with '^' and quotes with '^^'
            $SearchString = $SearchString.Replace($quotedString, $quotedString.Replace(" ", "^").Replace('"', "^^"));
        }
 
        # Split the $SearchString based on one or more blank spaces
        $stringComponents = $SearchString.Split($Delimiter,[StringSplitOptions]'RemoveEmptyEntries'); 
        for ($i = 0; $i -le $stringComponents.Length; $i++) {
            # The standard Powershell CompareTo method is case-sensitive
            if ([string]::Compare($stringComponents[$i], $PropertyName, $True) -eq 0) {
                # Check that we're not over the array boundary
                if ($i+1 -le $stringComponents.Length) {
                    # Restore any escaped quotation marks and spaces
                    $propertyValue = $stringComponents[$i+1].Replace("^^", '"').Replace("^", " ");
                    # Remove quotes
                    if ($RemoveQuotes) { $propertyValue = $propertyValue.Trim('"'); }
                    # Replace escaped quotes
                    if ($ReplaceEscapedQuotes) { return $propertyValue.Replace('�','"'); }
                    else { return $propertyValue.Replace('�','\"'); }
                }
            }
        }
        # If nothing has been found or we're over the array boundary, return the default value
        return $Default;
    }
}


<#
.SYNOPSIS
   Get an array of properies from a delimited string.
.DESCRIPTION
   The Get-StringProperySplit cmdlet returns an array of space-separated 
   strings from the source string, accounting for quoted text and escaped
   quotations.

   A single string can be returned with the -Index parameter. This parameter
   shortcuts allows replacing calls like '(Get-StringPropertySplit
    -SearchString $Source)[3]' with 'Get-StringPropertySplit -SearchString
    $Source -Index 3'
.PARAMETER SearchString
   String to search for the specified property name.
.PARAMETER Delimter
   The delimiter string/char to use. Regular expressions are supported and defaults to ' +'.
.PARAMETER RemoveQuotes
    Removes quotes from returned property values if present.
.PARAMETER ReplaceEscapedQuotes
    Replaces escaped quotes (\") with quotes (") from the returned property values if present.
.PARAMETER Index
    Returns the [string] at the specified index rather than a [string[]]
.EXAMPLE
   Get-StringPropertySplit -SearchString $StringToSearch

   This command returns an array of strings for all space-delimited values in the
   $StringToSearch variable, accounting for quoted strings and escaped quotes.
.EXAMPLE
   Get-StringPropertySplit $StringToSearch -RemoveQuotes

   This command returns an array of strings for all space-delimited values in the
   $StringToSearch variable, accounting for quoted strings and escaped quotes. All
   quotation marks are removed from quoted strings.
.EXAMPLE
   Get-StringPropertySplit $StringToSearch -ReplaceEscapedQuotes -Index 2

   This command returns a single string for the space-delimied value at array
   index 2 (third element), accounting for quoted strings and escaped quotes. The
   return string will have all escaped quotes '\"' replaced with '"'.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Get-StringPropertySplit {
    [CmdletBinding(HelpUri = 'http://virtualengine.co.uk/2014/searching-for-string-properties-with-powershell/')]
    [OutputType([String[]])]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString,
        # String delimiter, default to one or more spaces
        [Parameter(ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [string] $Delimiter = ' +',
        # Remove quotes from quoted strings
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $RemoveQuotes,
        # Replace escaped quotes with quotes
        [Parameter(ValueFromPipelineByPropertyName=$true)] [Switch] $ReplaceEscapedQuotes,
        # Return the specified index
        [Parameter(ValueFromPipelineByPropertyName=$true)] [int] $Index = -1
    )

    Process {
        # First replace escaped quotes with '�'
        $SearchString = $SearchString.Replace('\"', '�');
        
        while ($SearchString.Contains('"')) {
            # Store the right-hand side temporarily, skipping the first quote
            $searchStringRight = $SearchString.Substring($SearchString.IndexOf('"') +1);
            # Extract the quoted text from the original string
            $quotedString = $SearchString.Substring($SearchString.IndexOf('"'), $searchStringRight.IndexOf('"') +2);
            # Replace the quoted text, replacing spaces with '^' and quotes with '^^'
            $SearchString = $SearchString.Replace($quotedString, $quotedString.Replace(' ', '^').Replace('"', '^^'));
        }

        $stringArray = $SearchString.Split($Delimiter,[StringSplitOptions]'RemoveEmptyEntries'); 
        # Replace all escaped characters
        for ($i = 0; $i -lt $StringArray.Length; $i++) { 
            $stringArray[$i] = $stringArray[$i].Replace('^^', '"').Replace('^', ' ');
            # Remove quotes
            if ($RemoveQuotes) { $stringArray[$i] = $stringArray[$i].Trim('"'); }
            # Replace escaped quotes
            if ($ReplaceEscapedQuotes) { $stringArray[$i] = $stringArray[$i].Replace('�','"'); }
            else { $stringArray[$i] = $stringArray[$i].Replace('�','\"'); }
        }

        if ($Index -ne -1) { return $stringArray[$Index]; }
        else { return $stringArray; }
    }
}

<#
.SYNOPSIS
   Gets the NetScaler expression from the specified string
.DESCRIPTION
   This cmdlet returns a NetScaler expression that is escaped with 'q/' and is
   terminated by '/'. If a NetScaler expression is not found, $null is returned.
.PARAMETER SearchString
   String to search for the specified property name.
EXAMPLE
   Get-StringProperty -SearchString $StringToSearch

   This command searches the $StringToSearch variable for the presence of the a
   NetScaler expression, i.e. q/ .. /
.EXAMPLE
   Get-StringProperty $StringToSearch "-property3" "Not found"

   This command searches the $StringToSearch variable for the presence of the property
   "-property3" and returns its value, if found. If the property name is not found,
   the "Not Found" string will be returned.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Get-NetScalerExpression {
    [CmdletBinding()]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString
    )

    Process {
        $searchStringLeftPosition = $SearchString.IndexOf('q/');
        if ($searchStringLeftPosition -eq -1) { return $null; }
        $SearchString = $SearchString.Replace('q/', '��');

        $searchStringRightPosition = $SearchString.IndexOf('/');
        if ($searchStringRightPosition -eq -1) { return $null; }
        $SearchString = $SearchString.Replace('/', '�');

        $NetScalerExpression = $SearchString.Substring($searchStringLeftPosition, (($searchStringRightPosition +1)- $searchStringLeftPosition));
        return $NetScalerExpression.Replace('��','q/').Replace('�','/');
    }
}

<#
.SYNOPSIS
   Test for a named property value in a string.
.DESCRIPTION
   Tests for the presence of a property value in a string and returns a boolean
   value. For example, if the specified SearchString contained "-property1
   -property2 <value2>�, searching for "-Property1" or "-Property2" would return
   $true, but searching for "-Property3" would return $false
.PARAMETER SearchString
   String to search for the specified property name.
.PARAMETER PropertyName
   The property name to search the SearchString for.
.EXAMPLE
   Test-StringProperty -SearchString $StringToSearch -PropertyName "-property1"

   This command searches the $StringToSearch variable for the presence of the property
   "-property1". If the property name is found it returns $true. If the property name
   is not found, it will return $false.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringProperty {
    [CmdletBinding(HelpUri = 'http://virtualengine.co.uk/2014/searching-for-string-properties-with-powershell/')]
    [OutputType([bool])]
    Param (
        # String to search
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [ValidateNotNullOrEmpty()] [Alias("Search")] [string] $SearchString,
        # String property name to search for
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
        [ValidateNotNullOrEmpty()] [Alias("Name","Property")] [string] $PropertyName
    )

    Process {
        # Split the $SearchString based on one or more blank spaces
        $stringComponents = $SearchString.Split(' +',[StringSplitOptions]'RemoveEmptyEntries'); 
        for ($i = 0; $i -le $stringComponents.Length; $i++) {
            # The standard Powershell CompareTo method is case-sensitive
            if ([string]::Compare($stringComponents[$i], $PropertyName, $True) -eq 0) { return $true; }
        }
        # If nothing has been found or we're over the array boundary, return the default value
        return $false;
    }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Yes
    ($true) or No ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringPropertyYesNo([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "Yes"; }
    else { return "No"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string does not exist and returns
    either Yes ($false) or No ($true)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-NotStringPropertyYesNo([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "Yes"; }
    else { return "No"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Enabled
    ($true) or Disabled ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringPropertyEnabledDisabled([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "Enabled"; }
    else { return "Disabled"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Disabled
    ($true) or Enabled ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-NotStringPropertyEnabledDisabled([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "Enabled"; }
    else { return "Disabled"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either On
    ($true) or Off ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-StringPropertyOnOff([string]$SearchString, [string]$SearchProperty)
{
    if (Test-StringProperty $SearchString $SearchProperty) { return "On"; }
    else { return "Off"; }
}

<#
.SYNOPSIS
    Tests whether a named property in a string exists and returns either Off
    ($true) or On ($false)
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Test-NotStringPropertyOnOff([string]$SearchString, [string]$SearchProperty)
{
    if (-not (Test-StringProperty $SearchString $SearchProperty)) { return "On"; }
    else { return "Off"; }
}

<#
.SYNOPSIS
    Returns all strings that include the specified string $PropertyName parameter from 
    the array of string values passed into the $SearchString parameter.
.NOTES
   Author - Iain Brighton - @iainbrighton, iain.brighton@virtualengine.co.uk
#>

function Get-StringWithProperty {
    [CmdletBinding(DefaultParameterSetName='PropertyName')]
    [OutputType([string[]])]
    Param (
        [Parameter(Mandatory=$true, Position=0)] [ValidateNotNullOrEmpty()] [string[]] $SearchString,
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='PropertyName')] [ValidateNotNullOrEmpty()] [string] $PropertyName,
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='Like')] [ValidateNotNullOrEmpty()] [string] $Like
    )

    Begin {
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
        ## Check that we have a wildcard, somewhere
        if ($PSCmdlet.ParameterSetName -eq 'Like') {
            if (!$Like.Contains('*')) {
                ## If not do we need to add '*' or ' *'?
                if ($Like.EndsWith(' ')) {
                    Write-Warning "Get-StringWithProperty: No wildcard specified and '*' was appended.";
                    $Like += "*";
                } else {
                    Write-Warning "Get-StringWithProperty: No wildcard specified and ' *' was appended.";
                    $Like += " *";
                } # end if
            }
        }
	}

    Process {
        $MatchingStrings = @();
        foreach ($String in $SearchString) {
            switch ($PSCmdlet.ParameterSetName) {
                'PropertyName' {
                    if (Test-StringProperty -SearchString $String -PropertyName $PropertyName) {
                        $MatchingStrings += $String;
                    } # end if
                } # end propertyname
                'Like' {
                    if ($String -like $Like) {
                        $MatchingStrings += $String;
                    } #end if
                } # end like
            } # end switch
        } #end foreach
        return ,$MatchingStrings;
    } # end process
}

#endregion NetScaler Documentation Functions

#region NetScaler documentation pre-requisites

$SourceFileName = "ns.conf";

## Iain Brighton - Try and resolve the ns.conf file in the current working directory
if(Test-Path (Join-Path ((Get-Location).ProviderPath) $SourceFileName)) 
{
	$SourceFile = Join-Path ((Get-Location).ProviderPath) $SourceFileName; 
}
else 
{
	## Otherwise try the script's directory
	if(Test-Path (Join-Path (Split-Path $MyInvocation.MyCommand.Path) $SourceFileName)) 
	{
		$SourceFile = Join-Path (Split-Path $MyInvocation.MyCommand.Path) $SourceFileName; 
	}
	else 
	{
		throw "Cannot locate a NetScaler ns.conf file in either the working or script directory."; 
	}
}

#added by Carl Webster 24-May-2014
If(!$?)
{
	Write-Error "`n`n`t`tCannot locate a NetScaler ns.conf file in either the working or script directory.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

Write-Verbose "$(Get-Date): NetScaler file : $SourceFile"

## We read the file in once as each Get-Content call goes to disk and also creates a new string[]
$File = Get-Content $SourceFile

#added by Carl Webster 24-May-2014
If(!$? -or $Null -eq $File)
{
	Write-Error "`n`n`t`tUnable to read the NetScaler ns.conf file.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}
#endregion NetScaler documentation pre-requisites

#region NetScaler Create Collections

## Create smart and smaller collections for faster processing of the script.

$Add = Get-StringWithProperty -SearchString $File -Like 'add *';
$Set = Get-StringWithProperty -SearchString $File -Like 'set *';
$Bind = Get-StringWithProperty -SearchString $File -Like 'bind *';
$Enable = Get-StringWithProperty -SearchString $File -Like 'enable *';
$SetNS = Get-StringWithProperty -SearchString $Set -Like 'set ns *';
$SetNSTCPPARAM = Get-StringWithProperty -SearchString $SetNS -Like 'set ns tcpParam *';
$SetVpnParameter = Get-StringWithProperty -SearchString $Set -Like 'set vpn parameter *';
$SETAAAPREAUTH = Get-StringWithProperty -SearchString $Set -Like 'set aaa preauthenticationparameter *';
$AddCSPOLICY = Get-StringWithProperty -SearchString $Add -Like 'add cs policy *';
$ContentSwitches = Get-StringWithProperty -SearchString $Add -Like 'add cs vserver *';
$ContentSwitchBind = Get-StringWithProperty -SearchString $Bind -Like 'bind cs vserver *';
$CACHEREDIRS = Get-StringWithProperty -SearchString $Add -Like 'add cr vserver *';
$LoadBalancers = Get-StringWithProperty -SearchString $Add -Like 'add lb vserver *';
$LoadBalancerBind = Get-StringWithProperty -SearchString $Bind -Like 'bind lb vserver *';
$ServiceGroups = Get-StringWithProperty -SearchString $Add -Like 'add servicegroup *';
$ServiceGroupBind = Get-StringWithProperty -SearchString $Bind -Like 'bind servicegroup *';
$ServiceS = Get-StringWithProperty -SearchString $Add -Like 'add service *';
$ServiceBind = Get-StringWithProperty -SearchString $Bind -Like 'bind service *';
$Servers = Get-StringWithProperty -SearchString $Add -Like 'add server *';
$MONITORS = Get-StringWithProperty -SearchString $Add -Like 'add lb monitor *';
$NICS = Get-StringWithProperty -SearchString $Set -Like 'set interface *';
$CHANNELS = Get-StringWithProperty -SearchString $Set -Like 'set channel *';
$SIMPLEACLS = Get-StringWithProperty -SearchString $Add -Like 'add ns simpleacl *';
$IPList = Get-StringWithProperty -SearchString $Add -Like 'add ns ip *';
$VLANS = Get-StringWithProperty -SearchString $Add -Like 'add vlan *';
$VLANSBIND = Get-StringWithProperty -SearchString $BIND -Like 'bind vlan *';
$ROUTES = Get-StringWithProperty -SearchString $Add -Like 'add route *';
$AccessGateways = Get-StringWithProperty -SearchString $Add -Like 'add vpn vserver *';
$DNSNAMESERVERS = Get-StringWithProperty -SearchString $Add -Like 'add dns nameServer *';
$DNSRECORDCONFIGS = Get-StringWithProperty -SearchString $Add -Like 'add dns addRec *';
$CERTS = Get-StringWithProperty -SearchString $Add -Like 'add ssl certKey *';
$CERTBINDS = Get-StringWithProperty -SearchString $Bind -Like 'bind ssl vserver *';
$AUTHLDAPACTS = Get-StringWithProperty -SearchString $Add -Like 'add authentication ldapAction*';
$AUTHLDAPPOLS = Get-StringWithProperty -SearchString $Add -Like 'add authentication ldapPolicy*';
$AUTHRADIUSS = Get-StringWithProperty -SearchString $Add -Like 'add authentication radiusAction*';
$AUTHRADPOLS = Get-StringWithProperty -SearchString $Add -Like 'add authentication radiusPolicy*';
$AUTHGRPS = Get-StringWithProperty -SearchString $Add -Like 'add system group *';
$AUTHLOCS = Get-StringWithProperty -SearchString $Add -Like 'add system user *';
$AUTHLOCUSERS = Get-StringWithProperty -SearchString $Add -Like 'add authentication localPolicy *';
$CAGSESSIONPOLS = Get-StringWithProperty -SearchString $ADD -Like "add vpn sessionPolicy *";
$CAGSESSIONACTS = Get-StringWithProperty -SearchString $ADD -Like "add vpn sessionAction *";
$BINDVPNVSERVER = Get-StringWithProperty -SearchString $BIND -Like "bind vpn vserver *";
$CAGURLPOLS = Get-StringWithProperty -SearchString $ADD -Like "add vpn url *";

#new stuff added by Webster
$AdminPartitions = Get-StringWithProperty -SearchString $Add -Like 'add ns partition *';
$SSLProfiles = Get-StringWithProperty -SearchString $Add -Like 'add ssl profile *';
#endregion NetScaler Create Collections

#region NetScaler chaptercounters
$Chapters = 32
$Chapter = 0
#endregion NetScaler chaptercounters

#region NetScaler feature state
##Getting Feature states for usage later on and performance enhancements by not running parts of the script when feature is disabled
$Enable | ForEach-Object { 
    if ($_ -like 'enable ns feature *') {
        If ($_.Contains("WL") -eq "True") {$FEATWL = "Enabled"} Else {$FEATWL = "Disabled"}
        If ($_.Contains(" SP ") -eq "True") {$FEATSP = "Enabled"} Else {$FEATSP = "Disabled"}
        If ($_.Contains("LB") -eq "True") {$FEATLB = "Enabled"} Else {$FEATLB = "Disabled"}
        If ($_.Contains("CS") -eq "True") {$FEATCS = "Enabled"} Else {$FEATCS = "Disabled"}
        If ($_.Contains("CR") -eq "True") {$FEATCR = "Enabled"} Else {$FEATCR = "Disabled"}
        If ($_.Contains("SC") -eq "True") {$FEATSC = "Enabled"} Else {$FEATSC = "Disabled"}
        If ($_.Contains("CMP") -eq "True") {$FEATCMP = "Enabled"} Else {$FEATCMP = "Disabled"}
        If ($_.Contains("PQ") -eq "True") {$FEATPQ = "Enabled"} Else {$FEATPQ = "Disabled"}
        If ($_.Contains("SSL") -eq "True") {$FEATSSL = "Enabled"} Else {$FEATSSL = "Disabled"}
        If ($_.Contains("GSLB") -eq "True") {$FEATGSLB = "Enabled"} Else {$FEATGSLB = "Disabled"}
        If ($_.Contains("HDOSP") -eq "True") {$FEATHDSOP = "Enabled"} Else {$FEATHDOSP = "Disabled"}
        If ($_.Contains("CF") -eq "True") {$FEATCF = "Enabled"} Else {$FEATCF = "Disabled"}
        If ($_.Contains("IC") -eq "True") {$FEATIC = "Enabled"} Else {$FEATIC = "Disabled"}
        If ($_.Contains("SSLVPN") -eq "True") {$FEATSSLVPN = "Enabled"} Else {$FEATSSLVPN = "Disabled"}
        If ($_.Contains("AAA") -eq "True") {$FEATAAA = "Enabled"} Else {$FEATAAA = "Disabled"}
        If ($_.Contains("OSPF") -eq "True") {$FEATOSPF = "Enabled"} Else {$FEATOSPF = "Disabled"}
        If ($_.Contains("RIP") -eq "True") {$FEATRIP = "Enabled"} Else {$FEATRIP = "Disabled"}
        If ($_.Contains("BGP") -eq "True") {$FEATBGP = "Enabled"} Else {$FEATBGP = "Disabled"}
        If ($_.Contains("REWRITE") -eq "True") {$FEATREWRITE = "Enabled"} Else {$FEATREWRITE = "Disabled"}
        If ($_.Contains("IPv6PT") -eq "True") {$FEATIPv6PT = "Enabled"} Else {$FEATIPv6PT = "Disabled"}
        If ($_.Contains("AppFw") -eq "True") {$FEATAppFw = "Enabled"} Else {$FEATAppFw = "Disabled"}
        If ($_.Contains("RESPONDER") -eq "True") {$FEATRESPONDER = "Enabled"} Else {$FEATRESPONDER = "Disabled"}
        If ($_.Contains("HTMLInjection") -eq "True") {$FEATHTMLInjection = "Enabled"} Else {$FEATHTMLInjection = "Disabled"}
        If ($_.Contains("push") -eq "True") {$FEATpush = "Enabled"} Else {$FEATpush = "Disabled"}
        If ($_.Contains("AppFlow") -eq "True") {$FEATAppFlow = "Enabled"} Else {$FEATAppFlow = "Disabled"}
        If ($_.Contains("CloudBridge") -eq "True") {$FEATCloudBridge = "Enabled"} Else {$FEATCloudBridge = "Disabled"}
        If ($_.Contains("ISIS") -eq "True") {$FEATISIS = "Enabled"} Else {$FEATISIS = "Disabled"}
        If ($_.Contains("CH") -eq "True") {$FEATCH = "Enabled"} Else {$FEATCH = "Disabled"}
        If ($_.Contains("AppQoE") -eq "True") {$FEATAppQoE = "Enabled"} Else {$FEATAppQoE = "Disabled"}
        If ($_.Contains("Vpath") -eq "True") {$FEATVpath = "Enabled"} Else {$FEATVpath = "Disabled"}
        }
    }
#endregion NetScaler feature state

#region NetScaler Version

## Get version and build
$File | ForEach-Object { 
   if ($_ -like '#NS*') {
      $Y = ($_ -replace '#NS', '').split()
      $Version = $($Y[0]) 
      $Build = $($Y[2])
    }
}  

## Set script test version
## WIP THIS WORKS ONLY WHEN REGIONAL SETTINGS DIGIT IS SET TO . :)
$ScriptVersion = 10.1
#endregion NetScaler Version

#region NetScaler System Information

#region Basics
WriteWordLine 2 0 "NetScaler Basic"

$SETNS | ForEach-Object {
    if ($_ -like 'set ns hostname *') {
        $NSHOSTNAMEPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'set ns' ,'') -RemoveQuotes;
        $NSHOSTNAME = $NSHOSTNAMEPropertyArray[1];
        }
    }

$SAVEDDATESTRING = Get-StringWithProperty -SearchString $File -Like '# Last *';
$SAVEDDATESTRING | ForEach-Object { 
    $SAVEDDATE = ($_ -Replace '# Last modified by `save config`,' ,'') ;
    }

$Params = $null
$Params = @{
    Hashtable = @{
        Name = $NSHOSTNAME
        Version = $Version;
        Build = $Build;
        Saveddate = $SAVEDDATE
    }
    Columns = "Name","Version","Build","Saveddate";
    Headers = "Host Name","Version","Build","Last Configuration Saved Date";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#endregion Basics

#region NetScaler IP
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler IP"

WriteWordLine 2 0 "NetScaler Management IP Address"

$SetNS | ForEach-Object {
   if ($_ -like 'set ns config -IPAddress *') {
        $Params = $null
        $NSIP = Get-StringProperty $_ "-IPAddress";
        $Params = @{
            Hashtable = @{
                NSIP = Get-StringProperty $_ "-IPAddress";
                Subnet = Get-StringProperty $_ "-netmask";
            }
            Columns = "NSIP","Subnet";
            Headers = "NetScaler IP Address","Subnet";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        }
    }

#endregion NetScaler IP

#region NetScaler Global HTTP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global HTTP Parameters"

WriteWordLine 2 0 "NetScaler Global HTTP Parameters"
$SetNS | ForEach-Object {
    if ($_ -like 'set ns param *') {
        $IP = Get-StringProperty $_ "-cookieversion" "0";
        } else { $IP = "0" }
    if ($_ -like 'set ns httpParam *') {
        $DROP = Test-StringPropertyOnOff $_ "-dropInvalReqs";
        } else { $DROP = "On" }
    }

$Params = $null
$Params = @{
    Hashtable = @{
        CookieVersion = $IP;
        Drop = $DROP;
    }
    Columns = "CookieVersion","Drop";
    Headers = "Cookie Version","HTTP Drop Invalid Request";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "

#endregion NetScaler Global HTTP Parameters

#region NetScaler Global TCP Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global TCP Parameters"

WriteWordLine 2 0 "NetScaler Global TCP Parameters"

$SETNS | ForEach-Object {
   if ($_ -like 'set ns tcpParam *') {
        $TCP = Test-StringPropertyEnabledDisabled $_ "-WS";
        $SACK = Test-StringPropertyEnabledDisabled $_ "-SACK";
        $NAGLE = Test-StringPropertyEnabledDisabled $_ "-nagle";
    } else {
        $TCP = "Disabled";
        $SACK = "Disabled";
        $NAGLE = "Disabled";
    }
}

$Params = $null
$Params = @{
    Hashtable = @{
        TCP = $TCP;
        SACK = $SACK;
        NAGLE = $NAGLE;
    }
    Columns = "TCP","SACK","NAGLE";
    Headers = "TCP Windows Scaling","Selective Acknowledgement","Use Nagle's Algorithm";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;

FindWordDocumentEnd;
WriteWordLine 0 0 " "
    
#endregion NetScaler Global TCP Parameters

#region NetScaler Global Diameter Parameters
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Global Diameter Parameter"

WriteWordLine 2 0 "NetScaler Global Diameter Parameters"
$SetNS | ForEach-Object {
   if ($_ -like 'set ns diameter *') {
        $Params = $null
        $Params = @{
            Hashtable = @{
                HOST = Get-StringProperty $_ "-identity" "NA";
                Realm = Get-StringProperty $_ "-realm" "NA";
                Close = Get-StringProperty $_ "-serverClosePropagation" "No";

            }
            Columns = "HOST","Realm","Close";
            Headers = "Host Identity","Realm","Server Close Propagation";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        }
    }

#endregion NetScaler Global Diameter Parameters

#region NetScaler Time Zone
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Time zone"
WriteWordLine 2 0 "NetScaler Time Zone"

$Setns | ForEach-Object {  
    if ($_ -like 'set ns param *') {
        $TIMEZONE = Get-StringProperty $_ "-timezone" "Coordinated Universal Time" -RemoveQuotes;
        } else {$TIMEZONE = "Coordinated Universal Time"; }
    }

$Params = $null
$Params = @{
    Hashtable = @{
        TimeZone = $TIMEZONE;
    }
    Columns = "TimeZone";
    Headers = "Time Zone";
    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
    AutoFit = $wdAutoFitContent;
}
$Table = AddWordTable @Params -NoGridLines;
FindWordDocumentEnd;
WriteWordLine 0 0 " "

#endregion NetScaler Time Zone

#region NetScaler Management vLAN
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler management vLAN"

WriteWordLine 2 0 "NetScaler Management vLAN"

$SETNSDIAMETER = Get-StringWithProperty -SearchString $SetNS -Like 'set ns config -nsvlan *';
if($SETNSDIAMETER.Length -le 0) { WriteWordLine 0 0 "No Management vLAN has been configured"} else {
    $SetNS | ForEach-Object {
       if ($_ -like 'set ns config -nsvlan *') {
            $Params = $null
            $Params = @{
                Hashtable = @{
                    ## IB - This table will only have 1 row so create the nested hashtable inline
                    ID = Get-StringProperty $_ "-nsvlan";
                    INTERFACE = Get-StringProperty $_ "-ifnum";
                    Tagged = Get-StringProperty $_ "-tagged";
                }
                Columns = "ID","INTERFACE","Tagged";
                Headers = "vLAN ID","Interface","Tagged";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
            }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
    }
WriteWordLine 0 0 " "
#endregion NetScaler Management vLAN

#region NetScaler High Availability
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters High Availability"

WriteWordLine 2 0 "NetScaler High Availability"

$ADDHANODE = Get-StringWithProperty -SearchString $Add -Like 'add HA node *';
if($ADDHANODE.Length -le 0) { WriteWordLine 0 0 "High Availability has not been configured"} else {
    
    $NSRPCH = $null
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $NSRPCH = @();
    
    $NSRPCH += @{
        NSNODE = $NSIP; #NSIP Variable set in chapter NetScaler IP Address
    }

    $ADDHANODE | ForEach-Object {
        $HAPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add HA node' ,'') -RemoveQuotes;
        $NSRPCH += @{
            NSNODE = $HAPropertyArray[1];
        }
    }

    if ($NSRPCH.Length -gt 1) {
        $Params = $null
        $Params = @{
            Hashtable = $NSRPCH;
            Columns = "NSNODE";
            Headers = "NetScaler Node";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
            }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
    } else {WriteWordLine 0 0 "High Availability has not been configured"}
}
WriteWordLine 0 0 " "

#endregion NetScaler High Availability

#region NetScaler Administration
$selection.InsertNewPage()
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Administration"
WriteWordLine 2 0 "NetScaler System Authentication"
WriteWordLine 0 0 " "

#region Local Administration Users
WriteWordLine 3 0 "NetScaler System Users"

$AUTHLOCH = $null    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AUTHLOCH = @();

$AUTHLOCH += @{
    LocalUser = "nsroot";
    }

foreach ($AUTHLOC in $AUTHLOCS) {
    ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
    $AUTHLOCPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLOC -Replace 'add system user' ,'') -RemoveQuotes;

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be able 400 characters wide!
    $AUTHLOCH += @{
            LocalUser = $AUTHLOCPropertyArray[0];
        }
    }
if ($AUTHLOCH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $AUTHLOCH;
        Columns = "LocalUser";
        Headers = "Local User";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }
WriteWordLine 0 0 " "
#endregion Authentication Local Administration Users

#region Authentication Local Administration Groups
WriteWordLine 3 0 "NetScaler System Groups"

if($AUTHGRPS.Length -le 0) { WriteWordLine 0 0 "No Local Group has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHGRPH = @();

    foreach ($AUTHGRP in $AUTHGRPS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHGRPPropertyArray = Get-StringPropertySplit -SearchString ($AUTHGRP -Replace 'add system' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHGRPH += @{
                LocalUser = $AUTHGRPPropertyArray[1];
            }
        }

        if ($AUTHGRPH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHGRPH;
                Columns = "LocalUser";
                Headers = "Local User";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 " "
#endregion Authentication Local Administration Groups

#endregion NetScaler Administration

#region NetScaler Admin Partitions
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Admin Partitions"

WriteWordLine 2 0 "NetScaler Admin Partitions"

if($AdminPartitions.Length -le 0) 
{
	WriteWordLine 0 0 "No Admin Partitions have been configured"
} 
else 
{
       
	## IB - Use an array of hashtable to store the rows
	[System.Collections.Hashtable[]] $APH = @();
	
	foreach ($AdminPartition in $AdminPartitions) {
		## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
		$AdminPartitionPropertyArray = Get-StringPropertySplit -SearchString ($AdminPartition -Replace 'add ns partition' ,'') -RemoveQuotes;
		$AdminPartitionDisplayNameWithQoutes = Get-StringProperty $AdminPartition "partition";
		
		## IB - Create parameters for the hashtable so that we can splat them otherwise the
		## IB - command will be able 400 characters wide!

		$APBindMatches = Get-StringWithProperty -SearchString $AdminPartitionsBind -Like "bind ns partition $AdminPartitionDisplayNameWithQoutes *";

        $vlanap = ""
		$APBindMatches | ForEach-Object {
			$vlanap = Get-StringProperty $_ "-vlan";
			}

		$APH += @{
			ID = Get-StringProperty $AdminPartition "-partitionid";
			APNAME = $AdminPartitionPropertyArray[0];
			vLAN = $vlanap;
			MinBand = Get-StringProperty $AdminPartition "-minBandwidth" "10240";
			MaxBand = Get-StringProperty $AdminPartition "-maxBandwidth" "10240";
			Maxconn = Get-StringProperty $AdminPartition "-maxConn" "1024";
			Maxmem = Get-StringProperty $AdminPartition "-maxMemLimit" "10";
			}
	}

	if ($APH.Length -gt 0) 
	{
		$Params = $null
		$Params = @{
			Hashtable = $APH;
			Columns = "ID","APNAME","vLAN","MinBand","MaxBand","Maxconn","Maxmem";
			Headers = "ID","Name","vLAN","Minimum Bandwidth","Maximum Bandwidth","Maximum Connections","Maximum Memory";
			Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
			AutoFit = $wdAutoFitContent;
			}
		$Table = AddWordTable @Params -NoGridLines;
		FindWordDocumentEnd;
		WriteWordLine 0 0 " "
	}
}
#endregion NetScaler Admin Partitions

#region NetScaler Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Features"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Features"

If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, features added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }
#region NetScaler Basic Features
WriteWordLine 2 0 "NetScaler Basic Features"

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AdvancedConfiguration = @(
    @{ Description = "Feature"; Value = "State" }
	@{ Description = "Application Firewall"; Value = $FEATAppFw }
	@{ Description = "Authentication, Authorization and Auditing"; Value = $FEATAAA }
    @{ Description = "Content Filter"; Value = $FEATCF }
    @{ Description = "Content Switching"; Value = $FEATCS }
    @{ Description = "HTTP Compression"; Value = $FEATCMP }
    @{ Description = "Integrated Caching"; Value = $FEATIC }
    @{ Description = "Load Balancing"; Value = $FEATLB }
    @{ Description = "NetScaler Gateway"; Value = $FEATSSLVPN }
    @{ Description = "Rewrite"; Value = $FEATRewrite }
    @{ Description = "SSL Offloading"; Value = $FEATSSL }
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AdvancedConfiguration;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines -List;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      
 
WriteWordLine 0 0 " "

#endregion NetScaler Basic Features

#region NetScaler Advanced Features
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Advanced Features"

WriteWordLine 2 0 "NetScaler Advanced Features"

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $AdvancedFeatures = @(
    @{ Description = "Feature"; Value = "State" }
	@{ Description = "Web Logging"; Value = $FEATWL }
    @{ Description = "Surge Protection"; Value = $FEATSP }
    @{ Description = "Cache Redirection"; Value = $FEATCR }
    @{ Description = "Sure Connect"; Value = $FEATSC }
    @{ Description = "Priority Queuing"; Value = $FEATPQ }
    @{ Description = "Global Server Load Balancing"; Value = $FEATGSLB }
    @{ Description = "Http DoS Protection"; Value = $FEATHDOSP }
    @{ Description = "Vpath"; Value = $FEATVpath }
    @{ Description = "Integrated Caching"; Value = $FEATIC }
    @{ Description = "OSPF Routing"; Value = $FEATOSPF }
	@{ Description = "RIP Routing"; Value = $FEATRIP }
    @{ Description = "BGP Routing"; Value = $FEATBGP }
    @{ Description = "IPv6 protocol translation "; Value = $FEATIPv6PT }
    @{ Description = "Responder"; Value = $FEATRESPONDER }
    @{ Description = "Edgesight Monitoring HTML Injection"; Value = $FEATHTMLInjection }
    @{ Description = "OSPF Routing"; Value = $FEATOSPF }
    @{ Description = "NetScaler Push"; Value = $FEATPUSH }
    @{ Description = "AppFlow"; Value = $FEATAppFlow }
    @{ Description = "CloudBridge"; Value = $FEATCloudBridge }
    @{ Description = "ISIS Routing"; Value = $FEATISIS }
    @{ Description = "CallHome"; Value = $FEATCH }
    @{ Description = "AppQoE"; Value = $FEATAppQoE }
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $AdvancedFeatures;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params -NoGridLines -List;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      
 
WriteWordLine 0 0 " "

#endregion NetScaler Advanced Features

#endregion NetScaler Features

#region NetScaler Modes
$selection.InsertNewPage()
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Modes"

WriteWordLine 1 0 "NetScaler Modes"

If ($Version -gt $ScriptVersion) {
    WriteWordLine 0 0 ""
    WriteWordLine 0 0 "Warning: You are using Citrix NetScaler version $Version, modes added since version $ScriptVersion will not be shown."
    WriteWordLine 0 0 ""
    }

$Enable | ForEach-Object {  
    if ($_ -like 'enable ns mode *') {
        If ($_.Contains("FR") -eq "True") {$FR = "Enabled"} Else {$FR = "Disabled"}
        If ($_.Contains("L2") -eq "True") {$L2 = "Enabled"} Else {$L2 = "Disabled"}
        If ($_.Contains("USIP") -eq "True") {$USIP = "Enabled"} Else {$USIP = "Disabled"}
        If ($_.Contains("CKA") -eq "True") {$CKA = "Enabled"} Else {$CKA = "Disabled"}
        If ($_.Contains("TCPB") -eq "True") {$TCPB = "Enabled"} Else {$TCPB = "Disabled"}
        If ($_.Contains("MBF") -eq "True") {$MBF = "Enabled"} Else {$MBF = "Disabled"}
        If ($_.Contains("Edge") -eq "True") {$Edge = "Enabled"} Else {$Edge = "Disabled"}
        If ($_.Contains("USNIP") -eq "True") {$USNIP = "Enabled"} Else {$USNIP = "Disabled"}
        If ($_.Contains("PMTUD") -eq "True") {$PMTUD = "Enabled"} Else {$PMTUD = "Disabled"}
        If ($_.Contains("SRADV") -eq "True") {$SRADV = "Enabled"} Else {$SRADV = "Disabled"}
        If ($_.Contains("DRADV") -eq "True") {$DRADV = "Enabled"} Else {$DRADV = "Disabled"}
        If ($_.Contains("IRADV") -eq "True") {$IRADV = "Enabled"} Else {$IRADV = "Disabled"}
        If ($_.Contains("SRADV6") -eq "True") {$SRADV6 = "Enabled"} Else {$SRADV6 = "Disabled"}
        If ($_.Contains("DRADV6") -eq "True") {$DRADV6 = "Enabled"} Else {$DRADV6 = "Disabled"}
        If ($_.Contains("BridgeBPDUs") -eq "True") {$BridgeBPDUs = "Enabled"} Else {$BridgeBPDUs = "Disabled"}
        If ($_.Contains("L3") -eq "True") {$L3 = "Enabled"} Else {$L3 = "Disabled"}

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ADVModes = @(
            @{ Description = "Mode"; Value = "State"}  
            @{ Description = "Fast Ramp"; Value = $FR}        
            @{ Description = "Layer 2 mode"; Value = $L2}        
            @{ Description = "Use Source IP"; Value = $USIP}        
            @{ Description = "Client SideKeep-alive"; Value = $CKA}        
            @{ Description = "TCP Buffering"; Value = $TCPB}        
            @{ Description = "MAC-based forwarding"; Value = $MBF}
            @{ Description = "Edge configuration"; Value = $Edge}        
            @{ Description = "Use Subnet IP"; Value = $USNIP}        
            @{ Description = "Use Layer 3 Mode"; Value = $L3}        
            @{ Description = "Path MTU Discovery"; Value = $PMTUD}        
            @{ Description = "Static Route Advertisement"; Value = $SRADV}        
            @{ Description = "Direct Route Advertisement"; Value = $DRADV}        
            @{ Description = "Intranet Route Advertisement"; Value = $IRADV}        
            @{ Description = "Ipv6 Static Route Advertisement"; Value = $SRADV6}        
            @{ Description = "Ipv6 Direct Route Advertisement"; Value = $DRADV6}        
            @{ Description = "Bridge BPDUs" ; Value = $BridgeBPDUs}        
        );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $ADVModes;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

        FindWordDocumentEnd;
        $TableRange = $Null
        $Table = $Null      
        WriteWordLine 0 0 " "
        }
    }

$selection.InsertNewPage()

#endregion NetScaler Modes

#region NetScaler Web Interface
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Web Interface"

WriteWordLine 1 0 "NetScaler Web Interface"

$WI = Get-StringWithProperty -SearchString $File -Like 'install wi package *';

if($WI.Length -le 0) { WriteWordLine 0 0 "Citrix Web Interface has not been installed"} else { WriteWordLine 0 0 "Citrix Web Interface has been installed"}

$selection.InsertNewPage()

#endregion NetScaler Web Interface

#endregion NetScaler System Information

#region NetScaler Monitoring
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitoring"

WriteWordLine 1 0 "NetScaler Monitoring"

WriteWordLine 2 0 "SNMP Community"

$SNMPCOMS = Get-StringWithProperty -SearchString $Add -Like 'add snmp community *';

if($SNMPCOMS.Length -le 0) { WriteWordLine 0 0 "No SNMP Community has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPCOMH = @();

    foreach ($SNMPCOM in $SNMPCOMS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $SNMPCOMPropertyArray = Get-StringPropertySplit -SearchString ($SNMPCOM -Replace 'add snmp community' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $SNMPCOMH += @{
                SNMPCommunity = $SNMPCOMPropertyArray[0];
                Permission = $SNMPCOMPropertyArray[1];
            }
        }
        if ($SNMPCOMH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SNMPCOMH;
                Columns = "SNMPCommunity","Permission";
                Headers = "SNMP Community","Permission";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
        }
    }
WriteWordLine 0 0 " "

WriteWordLine 2 0 "SNMP Manager"

$SNMPMANS = Get-StringWithProperty -SearchString $Add -Like 'add snmp manager *';

if($SNMPMANS.Length -le 0) { WriteWordLine 0 0 "No SNMP Manager has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPMANSH = @();

    foreach ($SNMPMAN in $SNMPMANS) {

        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $SNMPMANPropertyArray = Get-StringPropertySplit -SearchString ($SNMPMAN -Replace 'add snmp manager' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $SNMPMANSH += @{
                SNMPManager = $SNMPMANPropertyArray[0];
                Netmask = Get-StringProperty $SNMPMAN "-netmask" "255.255.255.255";
            }
        }
        if ($SNMPMANSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SNMPMANSH;
                Columns = "SNMPManager","Netmask";
                Headers = "SNMP Manager","Netmask";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
        }
    }
WriteWordLine 0 0 ""

WriteWordLine 2 0 "SNMP Alert"

$SNMPALERTS = Get-StringWithProperty -SearchString $Set -Like 'set snmp alarm *';

if($SNMPALERTS.Length -le 0) { WriteWordLine 0 0 "No SNMP Alert has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SNMPALERTSH = @();

    foreach ($SNMPALERT in $SNMPALERTS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $SNMPALERTPropertyArray = Get-StringPropertySplit -SearchString ($SNMPALERT -Replace 'set snmp alarm' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $SNMPALERTSH += @{
                Alarm = $SNMPALERTPropertyArray[0];
                State = Test-NotStringPropertyEnabledDisabled $SNMPALERT "-state";
                Time = Get-StringProperty $SNMPALERT "-time" "0";
                TimeOut = Get-StringProperty $SNMPALERT "-timeout" "NA";
            }
        }
        if ($SNMPALERTSH.Length -gt 0) {
            $Params = @{
                Hashtable = $SNMPALERTSH;
                Columns = "Alarm","State","Time","TimeOut";
                Headers = "NetScaler Alarm","State","Time","Time-Out";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 ""

$selection.InsertNewPage()

#endregion NetScaler Monitoring

#region networking

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Network Configuration"

WriteWordLine 1 0 "NetScaler Networking"

#region NetScaler IP addresses

WriteWordLine 2 0 "NetScaler IP addresses"
if($IPLIST.Length -le 0) { WriteWordLine 0 0 "No IP Address has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $IPADDRESSH = @();

    foreach ($IP in $IPLIST) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $IPPropertyArray = Get-StringPropertySplit -SearchString ($IP -Replace 'add ns ip ' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $IPADDRESSH += @{
                IPAddress = $IPPropertyArray[0];
                SubnetMask = $IPPropertyArray[1];
                TrafficDomain = Get-StringProperty $IP "-td" "0";
                Management = Test-StringPropertyEnabledDisabled $IP "-mgmtAccess";
                vServer = Test-NotStringPropertyEnabledDisabled $IP "-vServer";
                GUI = Test-NotStringPropertyEnabledDisabled $IP "-gui";
                SNMP = Test-NotStringPropertyEnabledDisabled $IP "-snmp";
                Telnet = Test-NotStringPropertyEnabledDisabled $IP "-telnet";
            }
        }

        if ($IPADDRESSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $IPADDRESSH;
                Columns = "IPAddress","SubnetMask","TrafficDomain","Management","vServer","GUI","SNMP","Telnet";
                Headers = "IP Address","Subnet Mask","Traffic Domain","Management","vServer","GUI","SNMP","Telnet";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler IP addresses

#region NetScaler Interfaces

WriteWordLine 2 0 "NetScaler Interfaces"

if($NICS.Length -le 0) { WriteWordLine 0 0 "No network interface has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $NICH = @();

    foreach ($NIC in $NICS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        
        $NICDisplayName = Get-StringProperty $NIC "interface" -RemoveQuotes;
        
        $NICH += @{
            InterfaceID = $NICDisplayName;
            InterfaceType = Get-StringProperty $NIC "-intftype" -RemoveQuotes;
            HAMonitoring = Test-NotStringPropertyOnOff $NIC "-haMonitor";
            State = Test-NotStringPropertyOnOff $NIC "-state";
            AutoNegotiate = Test-NotStringPropertyOnOff $NIC "-autoneg";
            Tag = Test-StringPropertyOnOff $NIC "-tagall";
            }
        }

        if ($NICH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $NICH;
                Columns = "InterfaceID","InterfaceType","HAMonitoring","State","AutoNegotiate","Tag";
                Headers = "Interface ID","Interface Type","HA Monitoring","State","Auto Negotiate","Tag All vLAN";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler Interfaces

#region NetScaler vLAN

WriteWordLine 2 0 "NetScaler vLANs"

if($VLANS.Length -le 0) { WriteWordLine 0 0 "No vLAN has been configured"} else {
    $vLANH = $null
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $vLANH = @();

    foreach ($VLAN in $VLANS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $VLANDisplayName = Get-StringProperty $VLAN "vlan" -RemoveQuotes;
        $VLANBINDMatches = $null 
        $VLANBIND = $null
        
        $VLANBINDMatches = Get-StringWithProperty -SearchString $VLANSBIND -Like "bind vlan $VLANDisplayName *";

        foreach ($VLANBIND in $VLANBINDMatches) {
            $INT1 = Get-StringProperty $VLANBIND "-ifnum";
            }

        $vLANH += @{
            vLANID = $VLANDisplayName;
            Interface1 = $INT1;
            }
        }

        if ($vLANH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $vLANH;
                Columns = "vLANID","Interface1","Interface2","Interface3","Interface4","Interface5";
                Headers = "vLAN ID","Interface","Interface","Interface","Interface","Interface";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler vLAN

#region NetScaler Network Channel
$selection.InsertNewPage()
WriteWordLine 2 0 "NetScaler Network Channel"

if($CHANNELS.Length -le 0) { WriteWordLine 0 0 "No network channel has been configured"} else {
    
    foreach ($CHANNEL in $CHANNELS) {  
        $CHANNELDisplayName = Get-StringProperty $CHANNEL "channel" -RemoveQuotes;
        
        WriteWordLine 3 0 "Network Channel $CHANNELDisplayName"
        $Params = $null
        $Params = @{
            Hashtable = @{
            CHANNEL = $CHANNELDisplayName;
            Alias = Get-StringProperty $CHANNEL "-ifalias" "Not Configured";
            HA = Test-NotStringPropertyOnOff $CHANNEL "haMonitor";
            State = Get-StringProperty $CHANNEL "-state" "Enabled";
            Speed = Get-StringProperty $CHANNEL "-speed" "Auto";
            Tagall = Test-StringPropertyOnOff $CHANNEL "-Tagall";
            MTU = Get-StringProperty $CHANNEL "-mtu" "1500";
            }
        Columns = "CHANNEL","Alias","HA","State","Speed","Tagall";
        Headers = "Channel","Alias","HA Monitoring","State","Speed","Tag all vLAN";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
    
        $NICH = $NULL
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $NICH = @();

        foreach ($NIC in $NICS) {
            $CHANNELNIC = Get-StringProperty $NIC "-ifnum";
            If ($CHANNELNIC -eq $CHANNELDisplayName) {
                ## IB - Create parameters for the hashtable so that we can splat them otherwise the
                ## IB - command will be able 400 characters wide!
                $NICDisplayName = Get-StringProperty $NIC "interface" -RemoveQuotes;
                $NICH += @{
                    InterfaceID = $NICDisplayName;
                    }
                }
            }

        if ($NICH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $NICH;
                Columns = "InterfaceID";
                Headers = "Interface ID";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        }

        if($VLANS.Length -le 0) { WriteWordLine 0 0 "No vLAN has been configured for this Network Channel"} else {
            $vLANH = $null
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $vLANH = @();

            foreach ($VLAN in $VLANSBIND) {
                $CHANNELVLAN = $null
                $CHANNELVLAN = Get-StringProperty $VLAN "-ifnum";
                If ($CHANNELVLAN -eq $CHANNELDisplayName) {      
                    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
                    ## IB - command will be able 400 characters wide!

                    $VLANDisplayName = Get-StringProperty $VLAN "vlan" -RemoveQuotes;
                    $VLANH += @{
                        VLANID = $VLANDisplayName;
                        }
                    }
                }

            if ($VLANH.Length -gt 0) {
                $Params = $null
                $Params = @{
                    Hashtable = $VLANH;
                    Columns = "VLANID";
                    Headers = "VLAN ID";
                    Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                    AutoFit = $wdAutoFitContent;
                    }
                $Table = AddWordTable @Params -NoGridLines;
                FindWordDocumentEnd;
                WriteWordLine 0 0 " "
                } else { WriteWordLine 0 0 "No vLAN has been configured for this Network Channel"}
            }
        }
    }

WriteWordLine 0 0 " "
$selection.InsertNewPage()
#endregion NetScaler Network Channel

#region routing table

WriteWordLine 2 0 "NetScaler Routing Table"

WriteWordLine 0 0 "The NetScaler documentation script only documents manually added route table entries."
WriteWordLine 0 0 " "

if($ROUTES.Length -le 0) { WriteWordLine 0 0 "No route has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $ROUTESH = @();

    foreach ($ROUTE in $ROUTES) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $ROUTEPropertyArray = Get-StringPropertySplit -SearchString ($ROUTE -Replace 'add route ' ,'') -RemoveQuotes;
        $ROUTESH += @{
            Network = $ROUTEPropertyArray[0];
            Subnet = $ROUTEPropertyArray[1];
            Gateway = $ROUTEPropertyArray[2];
            Distance = Get-StringProperty $ROUTE "-distance" "0" -RemoveQuotes;
            Weight = Get-StringProperty $ROUTE "-weight" "1" -RemoveQuotes;
            Cost = Get-StringProperty $ROUTE "-cost" "0" -RemoveQuotes;
            }
        }

        if ($ROUTESH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $ROUTESH;
                Columns = "Network","Subnet","Gateway","Distance","Weight","Cost";
                Headers = "Network","Subnet","Gateway","Distance","Weight","Cost";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion routing table

#region NetScaler DNS Configuration
$selection.InsertNewPage()
WriteWordLine 2 0 "NetScaler DNS Configuration"

#region dns records

WriteWordLine 3 0 "NetScaler DNS Name Servers"
if($DNSNAMESERVERS.Length -le 0) { WriteWordLine 0 0 "No DNS Name Server has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSNAMESERVERH = @();

    foreach ($DNSNAMESERVER in $DNSNAMESERVERS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $DNSSERVERDisplayName = $null
        $DNSSERVERDisplayName = Get-StringProperty $DNSNAMESERVER "nameServer" -RemoveQuotes;
        $DNSNAMESERVERH += @{
            DNSServer = $DNSSERVERDisplayName;
            State = Get-StringProperty $DNSNAMESERVER "-state" "Enabled" -RemoveQuotes;
            Prot = Get-StringProperty $DNSNAMESERVER "-type" "UDP" -RemoveQuotes;
            }
        }

        if ($DNSNAMESERVERH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSNAMESERVERH;
                Columns = "DNSServer","State","Prot";
                Headers = "DNS Name Server","State","Protocol";;
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }
      
#endregion dns records

#region DNS Address Records

WriteWordLine 3 0 "NetScaler DNS Address Records"

if($DNSRECORDCONFIGS.Length -le 0) { WriteWordLine 0 0 "No DNS Address Record has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DNSRECORDCONFIGH = @();

    foreach ($DNSRECORDCONFIG in $DNSRECORDCONFIGS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $DNSRECORDCONFIGPropertyArray = Get-StringPropertySplit -SearchString ($DNSRECORDCONFIG -Replace 'add dns addRec ' ,'') -RemoveQuotes;
        $DNSRECORDCONFIGH += @{
            DNSRecord = $DNSRECORDCONFIGPropertyArray[0];
            IPAddress = $DNSRECORDCONFIGPropertyArray[1];
            TTL = Get-StringProperty $DNSRECORDCONFIG "-TTL";
            }
        }

        if ($DNSRECORDCONFIGH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $DNSRECORDCONFIGH;
                Columns = "DNSRecord","IPAddress","TTL";
                Headers = "DNS Record","IP Address","TTL";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion DNS Address Records

#endregion NetScaler DNS Configuration

#region NetScaler ACL
$selection.InsertNewPage()
WriteWordLine 2 0 "NetScaler ACL Configuration"

#region NetScaler Simple ACL IPv4

WriteWordLine 3 0 "NetScaler Simple ACL IPv4"

if($SIMPLEACLS.Length -le 0) { WriteWordLine 0 0 "No Simple ACL has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $SIMPLEACLSH = @();

    foreach ($SIMPLEACL in $SIMPLEACLS) {
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $SIMPLEACLPropertyArray = Get-StringPropertySplit -SearchString ($SIMPLEACL -Replace 'add ns simpleacl' ,'') -RemoveQuotes;

        $SIMPLEACLSH += @{
            ACLNAME = $SIMPLEACLPropertyArray[0];
            ACTION = $SIMPLEACLPropertyArray[1];
            SOURCEIP = Get-StringProperty $SIMPLEACL "-srcIP";
            DESTPORT = Get-StringProperty $SIMPLEACL "-destPort";
            PROT = Get-StringProperty $SIMPLEACL "-protocol";
            }
        }

        if ($SIMPLEACLSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SIMPLEACLSH;
                Columns = "ACLNAME","ACTION","SOURCEIP","DESTPORT","PROT";
                Headers = "ACL Name","Action","Source IP","Destination Port","Protocol";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

#endregion NetScaler Simple ACL IPv4

$selection.InsertNewPage()
#endregion NetScaler ACL

#endregion networking

#region NetScaler Traffic Domains
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Traffic Domains"

WriteWordLine 1 0 "NetScaler Traffic Domains"
##No function yet for routing table per TD

$TD = Get-StringWithProperty -SearchString $Add -Like 'add ns trafficDomain *';

if($TD.Length -le 0) { WriteWordLine 0 0 "No Traffic Domains have been configured"} else {
    WriteWordLine 0 0 "Only documents one assigned vLAN just like VLAN and Interface WIP"
    $TD | ForEach-Object {
        $Bind | ForEach-Object {  
            if ($_ -like 'bind ns trafficDomain *') {
                $vLAN = Get-StringProperty $_ "-vlan"
                }
            }
       
        $TDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add ns trafficDomain' ,'') -RemoveQuotes;        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $TDDisplayName = Get-StringProperty $_ "trafficDomain" -RemoveQuotes;
        $TDName = $TDDisplayName.Trim();
        
        WriteWordLine 2 0 "Traffic Domain $TDDisplayName"
        WriteWordLine 0 0 " "
        Write-Verbose "$(Get-Date): `tTraffic Domain $TDDisplayName"

        $TDDisplayName = Get-StringProperty $_ "trafficdomain" -RemoveQuotes;
        $TDName = $TDDisplayName.Trim();

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                ID = $TDPropertyArray[0];
                Alias = Get-StringProperty $_ "-aliasName";
                vLAN = $vLAN;
            }
            Columns = "ID","Alias","vLAN";
            Headers = "Traffic Domain ID","Traffic Domain Alias","Traffic Domain vLAN";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

        FindWordDocumentEnd;

        WriteWordLine 0 0 " "
        
        WriteWordLine 4 0 "Content Switch"        
    
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $CSTDH = @();

        $ContentSwitches | ForEach-Object {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $CSTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add cs vserver' ,'') -RemoveQuotes;
                $CSTDH = @{
                    ContentSwitch = $CSTDPropertyArray[0]
                }
            }
        }

        if ($CSTDH.Length -gt 0) {
            $Params = @{
                Hashtable = $CSTDH;
                Columns = "ContentSwitch";
                Headers = "Content Switch";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Content Switch been configured for this Traffic Domain"
            } # end if        
        
        WriteWordLine 4 0 "Load Balancer"      
  
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $LBTDH = @();

        $LoadBalancers | ForEach-Object {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $LBTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add lb vserver' ,'') -RemoveQuotes;
                $LBTDH += @{
                    LBTD = $LBTDPropertyArray[0];
                }
            }
        }
        
        if ($LBTDH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $LBTDH;
                Columns = "LBTD";
                Headers = "Load Balancer";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Load Balancer been configured for this Traffic Domain"
            } # end if

        WriteWordLine 4 0 "Services"      
    
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $SVCTDH = @();

        $Services | ForEach-Object {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $SVCTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add service' ,'') -RemoveQuotes;
                $SVCTDH += @{
                    SVCTD = $SVCTDPropertyArray[0];
                }
            }
        }
        
        if ($SVCTDH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SVCTDH;
                Columns = "SVCTD";
                Headers = "Service";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Service has been configured for this Traffic Domain"
            } # end if
  
        WriteWordLine 4 0 "Servers"      
  
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $SVRTDH = @();

        $Servers | ForEach-Object {
            if ((Get-StringProperty $_ "-td") -eq $TDPropertyArray[0]) {
                $SVRTDPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add server' ,'') -RemoveQuotes;
                $SVRTDH += @{
                    SVRTD = $SVRTDPropertyArray[0];
                }
            }
        }
        
        if ($SVRTDH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $SVRTDH;
                Columns = "SVRTD";
                Headers = "Server";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
        } else {
                WriteWordLine 0 0 "No Server has been configured for this Traffic Domain"
            } # end if        

        $selection.InsertNewPage()
        }
    }
    
#endregion NetScaler Traffic Domains

#region NetScaler Authentication
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Authentication"

$selection.InsertNewPage()

WriteWordLine 1 0 "NetScaler Authentication"

#region Local Users
WriteWordLine 2 0 "NetScaler Local Users"
if($AUTHLOCS.Length -le 0) { WriteWordLine 0 0 "No Local User has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHLOCUSRH = @();

    foreach ($AUTHLOCUSER in $AUTHLOCUSERS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHLOCUUSERPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLOCUSER -Replace 'add authentication localPolicy' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHLOCUSRH += @{
                LocalUser = $AUTHLOCUUSERPropertyArray[0];
                Expression = $AUTHLOCUUSERPropertyArray[1];
            }
        }
        if ($AUTHLOCUSRH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHLOCUSRH;
                Columns = "LocalUser","Expression";
                Headers = "Local User","Expression";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 " "
#endregion Authentication Local Users

#region Authentication LDAP Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler LDAP Authentication"
WriteWordLine 2 0 "NetScaler LDAP Policies"

if($AUTHLDAPPOLS.Length -le 0) { WriteWordLine 0 0 "No LDAP Policy has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHLDAPPOLH = @();

    foreach ($AUTHLDAPPOL in $AUTHLDAPPOLS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHLDAPPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLDAPPOL -Replace 'add authentication ldapPolicy' ,'') -RemoveQuotes;
                
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHLDAPPOLH += @{
                Policy = $AUTHLDAPPOLPropertyArray[0];
                Expression = $AUTHLDAPPOLPropertyArray[1];
                Action = $AUTHLDAPPOLPropertyArray[2];
            }
        }
        if ($AUTHLDAPPOLH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHLDAPPOLH;
                Columns = "Policy","Expression","Action";
                Headers = "LDAP Policy","Expression","LDAP Action";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;

            }
        }
WriteWordLine 0 0 " "

#endregion Authentication LDAP Policies

#region Authentication LDAP
WriteWordLine 2 0 "NetScaler LDAP authentication actions"

if($AUTHLDAPACTS.Length -le 0) { WriteWordLine 0 0 "No LDAP Authentication action has been configured"} else {
    $CurrentRowIndex = 0;
    foreach ($AUTHLDAP in $AUTHLDAPACTS) {
        
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHLDAPPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLDAP -Replace 'add authentication ldapAction' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $AUTHLDAPDisplayName = Get-StringProperty $AUTHLDAP "ldapAction" -RemoveQuotes;
        $AUTHLDAPName = $AUTHLDAPDisplayName.Trim();
    
        Write-Verbose "$(Get-Date): `tLDAP Authentication $CurrentRowIndex/$($AUTHLDAPS.Length) $AUTHLDAPDisplayName"     
        WriteWordLine 3 0 "LDAP Authentication action $AUTHLDAPDisplayName";

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $LDAPCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "LDAP Server IP"; Value = Get-StringProperty $AUTHLDAP "-serverIP"; }
            @{ Description = "LDAP Server Port"; Value = Get-StringProperty $AUTHLDAP "-serverPort" "389"; }
            @{ Description = "LDAP Server Time-Out"; Value = Get-StringProperty $AUTHLDAP "-authTimeout" "3"; }
            @{ Description = "Validate Certificate"; Value = Test-StringPropertyYesNo $AUTHLDAP "-validateServerCert"; }
            @{ Description = "LDAP Base OU"; Value = Get-StringProperty $AUTHLDAP "-ldapbase" -RemoveQuotes; }
            @{ Description = "LDAP Bind DN"; Value = Get-StringProperty $AUTHLDAP "-ldapBindDn" -RemoveQuotes; }
            @{ Description = "Login Name"; Value = Get-StringProperty $AUTHLDAP "-ldapLoginName"; }
            @{ Description = "Sub Attribute Name"; Value = Get-StringProperty $AUTHLDAP "-subAttributeName"; }
            @{ Description = "Security Type"; Value = Get-StringProperty $AUTHLDAP "-secType" "Default Setting"; }
            @{ Description = "Password Changes"; Value = Get-StringProperty $AUTHLDAP "-passwdChange" "Default Setting"; }
            @{ Description = "Search Filter"; Value = Get-StringProperty $AUTHLDAP "-searchFilter" "Not Configured" -RemoveQuotes;}
            @{ Description = "Group attribute name"; Value = Get-StringProperty $AUTHLDAP "-groupAttrName"; }
            @{ Description = "LDAP Single Sign On Attribute"; Value = Get-StringProperty $AUTHLDAP "-ssoNameAttribute" "Not Configured"; }
            @{ Description = "Authentication"; Value = Test-StringPropertyEnabledDisabled $AUTHLDAP "-authentication"; }
            @{ Description = "User Required"; Value = Test-StringPropertyYesNo $AUTHLDAP "-requireUser"; }
            @{ Description = "LDAP Referrals"; Value = Test-StringPropertyOnOff $AUTHLDAP "-followReferrals"; }
            @{ Description = "Nested Group Extraction"; Value = Test-NotStringPropertyOnOff $AUTHLDAP "-nestedGroupExtraction"; }
            @{ Description = "Maximum Nesting level"; Value = Get-StringProperty $AUTHLDAP "-maxNestingLevel" "Not Configured"; }
            );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $LDAPCONFIG;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null
        $selection.InsertNewPage()
    }
}
WriteWordLine 0 0 " "
#endregion Authentication LDAP

#region Authentication RADIUS Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Radius Authentication"
WriteWordLine 2 0 "NetScaler Radius Policies"

if($AUTHRADPOLS.Length -le 0) { WriteWordLine 0 0 "No Radius Policy has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $AUTHRADPOLH = @();

    foreach ($AUTHRADPOL in $AUTHRADPOLS) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHRADPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHRADPOL -Replace 'add authentication radiusPolicy' ,'') -RemoveQuotes;
      
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $AUTHRADPOLH += @{
                Policy = $AUTHRADPOLPropertyArray[0];
                Expression = $AUTHRADPOLPropertyArray[1];
                Action = $AUTHRADPOLPropertyArray[2];
            }
        }

        if ($AUTHRADPOLH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $AUTHRADPOLH;
                Columns = "Policy","Expression","Action";
                Headers = "Radius Policy","Expression","Radius Action";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            }
        }
WriteWordLine 0 0 " "
#endregion Authentication RADIUS Policies

#region Authentication RADIUS
WriteWordLine 2 0 "NetScaler RADIUS authentication action"

if($AUTHRADIUSS.Length -le 0) { WriteWordLine 0 0 "No RADIUS Authentication actions has been configured"} else {
    
    $CurrentRowIndex = 0;

    foreach ($AUTHRADIUS in $AUTHRADIUSS) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $AUTHRADIUSPropertyArray = Get-StringPropertySplit -SearchString ($AUTHRADIUS -Replace 'add authentication radiusAction' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $AUTHRADIUSDisplayName = Get-StringProperty $AUTHRADIUS "radiusAction" -RemoveQuotes;
        $AUTHRADIUSName = $AUTHRADIUSDisplayName.Trim();

        Write-Verbose "$(Get-Date): `tRADIUS Authentication $CurrentRowIndex/$($AUTHRADIUSS.Length) $AUTHRADIUSDisplayName"     
        WriteWordLine 3 0 "RADIUS Authentication action $AUTHRADIUSDisplayName";
        
        $RADIUSCONFIG = $null
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $RADIUSCONFIG = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "RADIUS Server IP"; Value = Get-StringProperty $AUTHRADIUS "-serverIP"; }
            @{ Description = "RADIUS Server Port"; Value = Get-StringProperty $AUTHRADIUS "-serverPort" "1812"; }
            @{ Description = "RADIUS Server Time-Out"; Value = Get-StringProperty $AUTHRADIUS "-authTimeout" "3"; }
            @{ Description = "Radius NAS IP"; Value = Test-StringPropertyEnabledDisabled $AUTHRADIUS "-radNASip"; }
            @{ Description = "Radius NAS ID"; Value = Get-StringProperty $AUTHRADIUS "-radNASid" "Not Configured"; }
            @{ Description = "Radius Vendor ID"; Value = Get-StringProperty $AUTHRADIUS "-radVendorID" "Not Configured"; }
            @{ Description = "Radius Attribute Type"; Value = Get-StringProperty $AUTHRADIUS "-radAttributeType" "Not Configured"; }
            @{ Description = "IP Vendor ID"; Value = Get-StringProperty $AUTHRADIUS "-ipVendorID" "Not Configured"; }
            @{ Description = "IP Attribute Type"; Value = Get-StringProperty $AUTHRADIUS "-ipAttributeType" "Not Configured"; }
            @{ Description = "Accounting"; Value = Test-StringPropertyOnOff $AUTHRADIUS "-accounting"; }
            @{ Description = "Password Vendor ID"; Value = Get-StringProperty $AUTHRADIUS "-pwdVendorID" "Not Configured"; }
            @{ Description = "Password Attribute Type"; Value = Get-StringProperty $AUTHRADIUS "-pwdAttributeType" "Not Configured"; }
            @{ Description = "Default Authentication Group"; Value = Get-StringProperty $AUTHRADIUS "-defaultAuthenticationGroup" "Not Configured"; }
            @{ Description = "Calling Station ID"; Value = Test-StringPropertyEnabledDisabled $AUTHRADIUS "-callingstationid"; }
            );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $RADIUSCONFIG;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null
        $selection.InsertNewPage()
    }
}
#endregion Authentication RADIUS

$selection.InsertNewPage()
#endregion NetScaler Authentication

#region NetScaler Certificates
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Certificates"

WriteWordLine 1 0 "NetScaler Certificates"

$CurrentRowIndex = 0;
if($CERTS.Length -le 0) { WriteWordLine 0 0 "No Certificate has been configured"} else {
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $CERTSH = @();
    
    $CERTS | ForEach-Object {
        $CurrentRowIndex++;
        $CERTPropertyArray = Get-StringPropertySplit -SearchString ($_ -Replace 'add ssl certKey' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $CERTSH += @{
            Certificate = $CERTPropertyArray[0]
            CertificateFile = Get-StringProperty $_ "-cert" -RemoveQuotes;
            CertificateKey = Get-StringProperty $_ "-key" -RemoveQuotes;
            Inform = Get-StringProperty $_ "-inform" "NA" -RemoveQuotes;
            }
        }
        if ($CERTSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $CERTSH;
                Columns = "Certificate","CertificateFile","CertificateKey","Inform";
                Headers = "Certificate","Certificate File","Certificate Key","Inform";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
    }

$selection.InsertNewPage()

#endregion NetScaler Certificates

#region traffic management

#region NetScaler Content Switches
$Chapter++

Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Content Switches"

WriteWordLine 1 0 "NetScaler Content Switches"

if($ContentSwitches.Length -le 0) { WriteWordLine 0 0 "No Content Switch has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($ContentSwitch in $ContentSwitches) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ContentSwitchPropertyArray = Get-StringPropertySplit -SearchString ($ContentSwitch -Replace 'add cs vserver' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $ContentSwitchDisplayNameWithQuotes = Get-StringProperty $ContentSwitch "vserver";
        $ContentSwitchDisplayName = Get-StringProperty $ContentSwitch "vserver" -RemoveQuotes;
        $ContentSwitchName = $ContentSwitchDisplayName.Trim();

        Write-Verbose "$(Get-Date): `tContent Switch $CurrentRowIndex/$($ContentSwitches.Length) $ContentSwitchDisplayName"     
        WriteWordLine 2 0 "Content Switch $ContentSwitchDisplayName";

        If (Test-StringProperty $ContentSwitch "-state") {$STATE = "Disabled"} else {$STATE = "Enabled"}

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                State = $STATE;
                Protocol = $ContentSwitchPropertyArray[1];
                Port = $ContentSwitchPropertyArray[3];
                IP = $ContentSwitchPropertyArray[2];
                TrafficDomain = Get-StringProperty $ContentSwitch "-td" "0 (Default)";
                CaseSensitive = Get-StringProperty $ContentSwitch "-caseSensitive";
                DownStateFlush = Get-StringProperty $ContentSwitch "-downStateFlush" "Least Connection";
                ClientTimeOut = Get-StringProperty $ContentSwitch "-cltTimeout" "NA";
            }
            Columns = "State","Protocol","Port","IP","TrafficDomain","CaseSensitive","DownStateFlush","ClientTimeOut";
            Headers = "State","Protocol","Port","IP","Traffic Domain","Case Sensitive","Down State Flush","Client Time-Out";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        WriteWordLine 3 0 "Policies"

        $ContentSwitchBindMatches = Get-StringWithProperty -SearchString $ContentSwitchBind -Like "bind cs vserver $ContentSwitchDisplayNameWithQuotes *";

        ## Check if we have any specific Content Switch bind matches
        if ($Null -eq $ContentSwitchBindMatches -or $ContentSwitchBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Policy has been configured for this Content Switch"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ContentSwitchPolicies = @();

            ## IB - Iterate over all Content Switch bindings (uses new function)
            foreach ($CSBind in $ContentSwitchBindMatches) {
                ## IB - Add each Content Switch binding with a policyName to the array
                if (Test-StringProperty -SearchString $CSBind -PropertyName "-policyName") {
                    ## Retrieve the service name from the Content Switch property array (position 3)

                    $AddCsPolicy | ForEach-Object {
                        if ($_ -like "add cs policy $(Get-StringProperty $CSBIND "-policyName") *") {
		                    $CSPOLRULE = Get-StringProperty $_ "-rule" -removequotes;
                          }
                        }

                    $ContentSwitchPolicies += @{
                        Policy = Get-StringProperty $CSBIND "-policyName"; 
                        "Load Balancer" = Get-StringProperty $CSBIND "-targetLBVserver";
                        Priority = Get-StringProperty $CSBIND "-priority";
                        Rule = $CSPOLRULE;
                        }
                    }
                } # end foreach

            if ($ContentSwitchPolicies.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!

                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $ContentSwitchPolicies;
                    Columns = "Policy","Load Balancer","Priority","Rule";
                    Headers = "Policy Name","Load Balancer","Priority","Rule";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No policy has been configured for this Content Switch"
            } # end if
        } #end if
        
        ##Table Redirect URL
        WriteWordLine 3 0 "Redirect URL"
            
        if (Test-StringProperty $ContentSwitch "-redirectURL") {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ContentSwitchRedirects = @(
                @{ RedirectURL = Get-StringProperty $ContentSwitch "-redirectURL" -RemoveQuotes; }
            );

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = @{
                    RedirectURL = Get-StringProperty $ContentSwitch "-redirectURL" -RemoveQuotes; 
                }
                Columns = "RedirectURL";
                Headers = "Redirect URL";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;

            FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No Redirect URL has been configured for this Content Switch"; }
            
            WriteWordLine 0 0 " "
            WriteWordLine 3 0 "Advanced Configuration"

            WriteWordLine 0 0 "Need to recheck if all options are correct"

             ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AdvancedConfiguration = @(                
                @{ Description = "Description"; Value = "Configuration"; }
                @{ Description = "Comment"; Value = Get-StringProperty $ContentSwitch "-comment" "No comment"; }
                @{ Description = "Apply AppFlow logging"; Value = Test-NotStringPropertyEnabledDisabled $ContentSwitch "-appflowLog"; }
                @{ Description = "Name of the TCP profile"; Value = Get-StringProperty $ContentSwitch "-tcpProfileName" "None"; }
                @{ Description = "Name of the HTTP profile"; Value = Get-StringProperty $ContentSwitch "-httpProfileName" "None"; }
                @{ Description = "Name of the NET profile"; Value = Get-StringProperty $ContentSwitch "-netProfile" "None"; }
                @{ Description = "Name of the DB profile"; Value = Get-StringProperty $ContentSwitch "-dbProfileName" "None"; }
                @{ Description = "Enable or disable user authentication"; Value = Test-StringPropertyOnOff $ContentSwitch "-Authentication"; }
                @{ Description = "Authentication virtual server FQDN"; Value = Get-StringProperty $ContentSwitch "-AuthenticationHost" "NA"; }
                @{ Description = "Name of the Authentication profile"; Value = Get-StringProperty $ContentSwitch "-authnProfile" "None"; }
                @{ Description = "Syntax expression identifying traffic"; Value = Test-StringPropertyOnOff $ContentSwitch "-Authentication"; }
                @{ Description = "Priority of the Listener Policy"; Value = Get-StringProperty $ContentSwitch "-AuthenticationHost" "NA"; }
                @{ Description = "Name of the backup virtual server"; Value = Get-StringProperty $ContentSwitch "-authnProfile" "None"; }
                @{ Description = "Enable state updates"; Value = Get-StringProperty $ContentSwitch "-Listenpolicy" "None"; }
                @{ Description = "Route requests to the cache server"; Value = Get-StringProperty $ContentSwitch "-Listenpriority" "101 (Maximum Value)"; }
                @{ Description = "Precedence to use for policies"; Value = Get-StringProperty $ContentSwitch "-backupVServer" "NA"; }
                @{ Description = "URL Case sensitive"; Value = Get-StringProperty $ContentSwitch "-timeout" "2 (Default Value)"; }
                @{ Description = "Type of spillover"; Value = Get-StringProperty $ContentSwitch "-persistenceBackup" "None"; }
                @{ Description = "Maintain source-IP based persistence"; Value = Get-StringProperty $ContentSwitch "-backupPersistenceTimeout" "2 (Default Value)"; }
                @{ Description = "Action if spillover is to take effect"; Value = Test-StringPropertyOnOff $ContentSwitch "-pq"; }
                @{ Description = "State of port rewrite HTTP redirect"; Value = Test-StringPropertyOnOff $ContentSwitch "-sc"; }
                @{ Description = "Continue forwarding to backup vServer"; Value = Test-StringPropertyOnOff $ContentSwitch "-rtspNat"; }
            );

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $AdvancedConfiguration;
                Columns = "Description","Value";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines -List;

		    FindWordDocumentEnd;
		    $TableRange = $Null
		    $Table = $Null     
        FindWordDocumentEnd;
        $selection.InsertNewPage()
        }
    }
$selection.InsertNewPage()

#endregion NetScaler Content Switches

#region NetScaler Cache Redirection
$Chapter++

Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Cache Redirection"

WriteWordLine 1 0 "NetScaler Cache Redirection"

if($CACHEREDIRS.Length -le 0) { WriteWordLine 0 0 "No Cache Redirection has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($CACHEREDIR in $CACHEREDIRS) {
        $CurrentRowIndex++;
        $CACHEREDIRPropertyArray = Get-StringPropertySplit -SearchString ($CACHEREDIR -Replace 'add cr vserver' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $CACHEREDIRDisplayNameWithQuotes = Get-StringProperty $CACHEREDIR "vserver";
        $CACHEREDIRDisplayName = Get-StringProperty $CACHEREDIR "vserver" -RemoveQuotes;

        Write-Verbose "$(Get-Date): `tCache Redirection $CurrentRowIndex/$($ContentSwitches.Length) $CACHEREDIRDisplayName"     
        WriteWordLine 2 0 "Cache Redirection $CACHEREDIRDisplayName";

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                NAME = $CACHEREDIRPropertyArray[0];
                PROT = $CACHEREDIRPropertyArray[1];
                IP = $CACHEREDIRPropertyArray[2];
                CACHETYPE = Get-StringProperty $CACHEREDIR "-cacheType" "0 (Default)";
                REDIRECT = Get-StringProperty $CACHEREDIR "-redirect";
                CLTTIEMOUT = Get-StringProperty $CACHEREDIR "-cltTimeout";
                DNSVSERVER = Get-StringProperty $CACHEREDIR "-dnsVserverName";
            }
            Columns = "NAME","PROT","IP","CACHETYPE","REDIRECT","CLTTIEMOUT","DNSVSERVER";
            Headers = "NAME","PROT","IP","CACHETYPE","REDIRECT","CLTTIEMOUT","DNSVSERVER";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        }
    }
$selection.InsertNewPage()

#endregion NetScaler Cache Redirection

#region NetScaler Load Balancers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Load Balancers"

WriteWordLine 1 0 "NetScaler Load Balancing"

if($LoadBalancers.Length -le 0) { WriteWordLine 0 0 "No Load Balancer has been configured"} else {
    ## IB - We no longer need to worrying about the number of columns and/or rows.
    ## IB - Need to create a counter of the current row index
    $CurrentRowIndex = 0;

    foreach ($LoadBalancer in $LoadBalancers) {

        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $LoadBalancerPropertyArray = Get-StringPropertySplit -SearchString ($LoadBalancer -Replace 'add lb vserver' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $LoadBalancerDisplayName = Get-StringProperty $LoadBalancer "vserver" -RemoveQuotes;
        $LoadBalancerDisplayNameWithQoutes = Get-StringProperty $LoadBalancer "vserver";
        $LoadBalancerName = $LoadBalancerDisplayName.Trim();

        Write-Verbose "$(Get-Date): `tLoad Balancer $CurrentRowIndex/$($LoadBalancers.Length) $LoadBalancerDisplayName"     
        WriteWordLine 2 0 "Load Balancer $LoadBalancerDisplayName";
        
        If (Test-StringProperty $LoadBalancer "-state") {$STATE = "Disabled"} else {$STATE = "Enabled"}

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                State = $STATE;
                Protocol = $LoadBalancerPropertyArray[1];
                Port = $LoadBalancerPropertyArray[3];
                IP = $LoadBalancerPropertyArray[2];
                Persistency = Get-StringProperty $LoadBalancer "-persistenceType";
                TrafficDomain = Get-StringProperty $LoadBalancer "-td" "0 (Default)";
                Method = Get-StringProperty $LoadBalancer "-lbmethod" "Least Connection";
                ClientTimeOut = Get-StringProperty $LoadBalancer "-cltTimeout" "NA";
            }
            Columns = "State","Protocol","Port","IP","Persistency","TrafficDomain","Method","ClientTimeOut";
            Headers = "State","Protocol","Port","IP","Persistency","Traffic Domain","Method","Client Time-Out";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;
        #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        ##Services Table
        WriteWordLine 3 0 "Service and Service Group"

        $LoadBalancerBindMatches = Get-StringWithProperty -SearchString $LoadbalancerBind -Like "bind lb vserver $LoadBalancerDisplayNameWithQoutes *";
        ## Check if we have any specific load balancer bind matches
        if ($Null -eq $LoadBalancerBindMatches -or $LoadBalancerBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Service (Group) has been configured for this Load Balancer"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerServices = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($LBBind in $LoadBalancerBindMatches) {
                ## IB - Add each load balancer binding with a policyName to the array
                if (-not (Test-StringProperty -SearchString $LBBind -PropertyName "-policyName")) {
                    ## Retrieve the service name from the load balancer property array (position 3)
                    $LoadBalancerServices += @{ Service = (Get-StringPropertySplit $LBBind)[4]; }
                }
            } # end foreach

            if ($LoadBalancerServices.Length -gt 0) {
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $LoadBalancerServices;
                    AutoFit = $wdAutoFitContent;
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No Service (Group) has been configured for this Load Balancer"
            } # end if
        }
        FindWordDocumentEnd;

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Policies"

        if ($LoadBalancerBind.Length -le 0) {
            WriteWordLine 0 0 "No Policy has been configured for this Load Balancer"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerPolicies = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($LBBind in (Get-StringWithProperty -SearchString $LoadBalancerBind -Like "bind lb vserver $LoadBalancerDisplayNameWithQoutes *")) {
                ## IB - Add each load balancer binding with a policyName to the array
                if (Test-StringProperty -SearchString $LBBind -PropertyName "-policyName") {
                    $LoadBalancerPolicies += @{
                        Name = Get-StringProperty $LBBind "-policyName" -RemoveQuotes;
                        Priority = Get-StringProperty $LBBind "-priority" "NA";
                        Type = Get-StringProperty $LBBind "-type" "NA";
                        Expression = Get-StringProperty $LBBind "-gotoPriorityExpression" "NA"; }
                } # end if
            } # end foreach

            if ($LoadBalancerPolicies.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!

                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $LoadBalancerPolicies;
                    Columns = "Name","Priority","Type","Expression";
                    Headers = "Policy Name","Priority","Policy Type","GoTo Expression";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No Policy has been configured for this Load Balancer"
            } # end if
        } #end if

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Redirect URL"

        if (Test-StringProperty $LoadBalancer "-redirectURL") {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $LoadBalancerRedirects = @(
                @{ RedirectURL = Get-StringProperty $LoadBalancer "-redirectURL" -RemoveQuotes; }
            );

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = @{
                    RedirectURL = Get-StringProperty $LoadBalancer "-redirectURL" -RemoveQuotes; 
                }
                Columns = "RedirectURL";
                Headers = "Redirect URL";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

            FindWordDocumentEnd;
        } else { WriteWordLine 0 0 "No Redirect URL has been configured for this Load Balancer"; }
        
        ##Advanced Configuration   
        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "Comment"; Value = Get-StringProperty $LoadBalancer "-comment" "No comment"; }
            @{ Description = "Apply AppFlow logging"; Value = Test-NotStringPropertyEnabledDisabled $LoadBalancer "-appflowLog"; }
            @{ Description = "Name of the TCP profile"; Value = Get-StringProperty $LoadBalancer "-tcpProfileName" "None"; }
            @{ Description = "Name of the HTTP profile"; Value = Get-StringProperty $LoadBalancer "-httpProfileName" "None"; }
            @{ Description = "Name of the NET profile"; Value = Get-StringProperty $LoadBalancer "-netProfile" "None"; }
            @{ Description = "Name of the DB profile"; Value = Get-StringProperty $LoadBalancer "-dbProfileName" "None"; }
            @{ Description = "Enable or disable user authentication"; Value = Test-StringPropertyOnOff $LoadBalancer "-Authentication"; }
            @{ Description = "Authentication virtual server FQDN"; Value = Get-StringProperty $LoadBalancer "-AuthenticationHost" "NA"; }
            @{ Description = "Authentication virtual server name"; Value = Get-StringProperty $LoadBalancer "-authnVsname" "NA"; }
            @{ Description = "Name of the Authentication profile"; Value = Get-StringProperty $LoadBalancer "-authnProfile" "None"; }
            @{ Description = "User authentication with HTTP 401"; Value = Test-StringPropertyOnOff $LoadBalancer "-authn401"; }
            @{ Description = "Syntax expression identifying traffic"; Value = Get-StringProperty $LoadBalancer "-Listenpolicy" "None"; }
            @{ Description = "Priority of the Listener Policy"; Value = Get-StringProperty $LoadBalancer "-Listenpriority" "101 (Maximum Value)"; }
            @{ Description = "Name of the backup virtual server"; Value = Get-StringProperty $LoadBalancer "-backupVServer" "NA"; }
            @{ Description = "Time period a persistence session"; Value = Get-StringProperty $LoadBalancer "-timeout" "2 (Default Value)"; }
            @{ Description = "Backup persistence type"; Value = Get-StringProperty $LoadBalancer "-persistenceBackup" "None"; }
            @{ Description = "Time period a backup persistence session"; Value = Get-StringProperty $LoadBalancer "-backupPersistenceTimeout" "2 (Default Value)"; }
            @{ Description = "Use priority queuing"; Value = Test-StringPropertyOnOff $LoadBalancer "-pq"; }
            @{ Description = "Use SureConnect"; Value = Test-StringPropertyOnOff $LoadBalancer "-sc"; }
            @{ Description = "Use network address translation"; Value = Test-StringPropertyOnOff $LoadBalancer "-rtspNat"; }
            @{ Description = "Redirection mode for load balancing"; Value = Get-StringProperty $LoadBalancer "-m" "IP Based"; }
            @{ Description = "Use Layer 2 parameter"; Value = Test-StringPropertyOnOff $LoadBalancer "-l2Conn"; }
            @{ Description = "TOS ID of the virtual server"; Value = Get-StringProperty $LoadBalancer "-tosId" "0 (Default)"; }
            @{ Description = "Expression against which traffic is evaluated"; Value = Get-StringProperty $LoadBalancer "-rule" "None"; }
            @{ Description = "Perform load balancing on a per-packet basis"; Value = Test-StringPropertyEnabledDisabled $LoadBalancer "-sessionless"; }
            @{ Description = "How the NetScaler appliance responds to ping requests"; Value = Get-StringProperty $LoadBalancer "-icmpVsrResponse" "NS_VSR_PASSIVE (Default)"; }
            @{ Description = "Route cacheable requests to a cache redirection server"; Value = Test-StringPropertyYesNo $LoadBalancer "-cacheable"; }
        );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $AdvancedConfiguration;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;
        ## IB - Set the header background and bold font
        #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null

        $selection.InsertNewPage()
    }
}

#endregion NetScaler Load Balancers

#region NetScaler Services
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Services"

FindWordDocumentEnd;

WriteWordLine 1 0 "NetScaler Services"

if($Services.Length -le 0) { WriteWordLine 0 0 "No Service has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Service in $Services) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ServicePropertyArray = Get-StringPropertySplit -SearchString ($Service -Replace 'add service' ,'') -RemoveQuotes;
        
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $ServiceDisplayNameWithQuotes = Get-StringProperty $Service "service";
        $ServiceDisplayName = Get-StringProperty $Service "service" -RemoveQuotes;
        $ServiceName = $ServiceDisplayName.Trim();
    
        Write-Verbose "$(Get-Date): `tService $CurrentRowIndex/$($Services.Length) $ServiceDisplayName"     
        WriteWordLine 2 0 "Service $ServiceDisplayName"

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                Server = $ServicePropertyArray[1];
                Protocol = $ServicePropertyArray[2];
                Port = $ServicePropertyArray[3];
                TD = Get-StringProperty $Service "-td" "0 (Default)";
                GSLB = Get-StringProperty $Service "-gslb" "NA";
                MaximumClients = Get-StringProperty $Service "-maxClient" "NA";
                MaximumRequests = Get-StringProperty $Service "-maxreq" "NA";
            }
            Columns = "Server","Protocol","Port","TD","GSLB","MaximumClients","MaximumRequests";
            Headers = "Server","Protocol","Port","Traffic Domain","GSLB","Maximum Clients","Maximum Requests";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        WriteWordLine 3 0 "Monitor"

        $ServiceBindMatches = Get-StringWithProperty -SearchString $ServiceBind -Like "bind service $ServiceDisplayNameWithQuotes *";
        ## Check if we have any specific Service bind matches
        if ($Null -eq $ServiceBindMatches -or $ServiceBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Monitor has been configured for this Service"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ServiceMonitors = @();

            ## IB - Iterate over all Service bindings (uses new function)
            foreach ($SVCBind in $ServiceBindMatches) {
                if (Test-StringProperty -SearchString $SVCBind -PropertyName "-monitorName") {
                    $ServiceMonitors += @{ Monitor = Get-StringProperty $SVCBIND "-monitorName"; }
                }
            } # end foreach

            if ($ServiceMonitors.Length -gt 0) {
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $ServiceMonitors;                   
                    AutoFit = $wdAutoFitContent;
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No Monitor has been configured for this Service"
            }
        } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "Clear text port"; Value = Get-StringProperty $Service "-clearTextPort" "NA" ; }
			@{ Description = "Cache Type"; Value = Get-StringProperty $Service "-cacheType" "NA" ; }
			@{ Description = "Maximum Client Requests"; Value = Get-StringProperty $Service "-maxClient" "4294967294 (Maximum Value)" ; }
			@{ Description = "Monitor health of this service"; Value = Test-NotStringPropertyYesNo $Service "-healthMonitor" ; }
			@{ Description = "Maximum Requests"; Value = Get-StringProperty $Service "-maxreq" "65535 (Maximum Value)" ; }
			@{ Description = "Use Transparent Cache"; Value = Test-StringPropertyYesNo $Service "-cacheable" ; }
			@{ Description = "Insert the Client IP header"; Value = Get-StringProperty $Service "-cip" "DISABLED"  ; }
			##@{ Description = "Name for the HTTP header"; Value = Get-StringProperty $Service "-cipHeader" "NA" ; }
			@{ Description = "Use Source IP"; Value = Test-NotStringPropertyYesNo $Service "-usip" ; }
            @{ Description = "Path Monitoring"; Value = Test-StringPropertyYesNo $Service "-pathMonitor" ; }
			@{ Description = "Individual Path monitoring"; Value = Test-StringPropertyYesNo $Service "-pathMonitorIndv" ; }
			@{ Description = "Use the proxy port"; Value = Test-StringPropertyYesNo $Service "-useproxyport" ; }
			@{ Description = "SureConnect"; Value = Test-StringPropertyOnOff $Service "-sc" ; }
			@{ Description = "Surge protection"; Value = Test-NotStringPropertyOnOff $Service "-sp" ; }
			@{ Description = "RTSP session ID mapping"; Value = Test-StringPropertyOnOff $Service "-rtspSessionidRemap" ; }
			@{ Description = "Client Time-Out"; Value = Get-StringProperty $Service "-cltTimeout" "31536000 (Maximum Value)" ; }
			@{ Description = "Server Time-Out"; Value = Get-StringProperty $Service "-svrTimeout" "3153600 (Maximum Value)" ; }
			@{ Description = "Unique identifier for the service"; Value = Get-StringProperty $Service "-CustomServerID" "None" -RemoveQuotes; }
			@{ Description = "The identifier for the service"; Value = Get-StringProperty $Service "-serverID" "None" ; }
			@{ Description = "Enable client keep-alive"; Value = Test-NotStringPropertyYesNo $Service "-CKA" ; }
			@{ Description = "Enable TCP buffering"; Value = Test-NotStringPropertyYesNo $Service "-TCPB" ; }
            @{ Description = "Enable compression"; Value = Test-StringPropertyYesNo $Service "-CMP" ; }
			@{ Description = "Maximum bandwidth, in Kbps"; Value = Get-StringProperty $Service "-maxBandwidth" "4294967287 (Maximum Value)" ; }
			@{ Description = "Use Layer 2 mode"; Value = Test-StringPropertyYesNo $Service "-accessDown" ; }
			@{ Description = "Sum of weights of the monitors"; Value = Get-StringProperty $Service "-monThreshold" "65535 (Maximum Value)" ; }
			@{ Description = "Initial state of the service"; Value = Test-NotStringPropertyEnabledDisabled $Service "-state" ; }
			@{ Description = "Perform delayed clean-up"; Value = Test-NotStringPropertyEnabledDisabled $Service "-downStateFlush" ; }
			@{ Description = "TCP profile"; Value = Get-StringProperty $Service "-tcppProfileName" "NA" ; }
			@{ Description = "HTTP profile"; Value = Get-StringProperty $Service "-httpProfileName" "NA" ; }
			@{ Description = "A numerical identifier"; Value = Get-StringProperty $Service "-hashId" "NA" ; }
			@{ Description = "Comment about the service"; Value = Get-StringProperty $Service "-comment" "NA"; }
			@{ Description = "Logging of AppFlow information"; Value = Test-NotStringPropertyEnabledDisabled $Service "-appflowLog" ; }
			@{ Description = "Network profile"; Value = Get-StringProperty $Service "-netProfile" "NA" ; }
        );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $AdvancedConfiguration;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null
        WriteWordLine 0 0 " "

        $selection.InsertNewPage() 
        }
   }

#endregion NetScaler Services

#region NetScaler Service Groups
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Service Groups"

FindWordDocumentEnd;

WriteWordLine 1 0 "NetScaler Service Groups"

if($ServiceGroups.Length -le 0) { WriteWordLine 0 0 "No Service Group has been configured"} else {
    $CurrentRowIndex = 0;

    foreach ($Servicegroup in $ServiceGroups) {
        $CurrentRowIndex++;
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ServicegroupPropertyArray = Get-StringPropertySplit -SearchString ($Servicegroup -Replace 'add servicegroup' ,'') -RemoveQuotes;
        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $ServicegroupDisplayNameWithQuotes = Get-StringProperty $Servicegroup "servicegroup";
        $ServicegroupDisplayName = Get-StringProperty $Servicegroup "servicegroup" -RemoveQuotes;
        $ServicegroupName = $ServicegroupDisplayName.Trim();
        
        Write-Verbose "$(Get-Date): `tService Group $CurrentRowIndex/$($ServiceGroups.Length) $ServicegroupDisplayName"     
        WriteWordLine 2 0 "Service Group $ServicegroupDisplayName"

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - This table will only have 1 row so create the nested hashtable inline
                Server = $ServicegroupPropertyArray[0];
                Protocol = $ServicegroupPropertyArray[1];
                Port = $ServicegroupPropertyArray[3];
                TD = Get-StringProperty $Servicegroup "-td" "0 (Default)";
                GSLB = Get-StringProperty $Servicegroup "-gslb" "NA";
                MaximumClients = Get-StringProperty $Servicegroup "-maxClient" "NA";
                MaximumRequests = Get-StringProperty $Servicegroup "-maxreq" "NA";
            }
            Columns = "Server","Protocol","Port","TD","GSLB","MaximumClients","MaximumRequests";
            Headers = "Server","Protocol","Port","Traffic Domain","GSLB","Maximum Clients","Maximum Requests";
            Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
            AutoFit = $wdAutoFitContent;
        }
        $Table = AddWordTable @Params -NoGridLines;

        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

        $ServicegroupBindMatches = Get-StringWithProperty -SearchString $ServiceGroupBind -Like "bind serviceGroup $ServicegroupDisplayNameWithQuotes *";

        WriteWordLine 3 0 "Servers"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $ServiceGroupServers = @();

        ## IB - Iterate over all Service bindings (uses new function)
        foreach ($Server in $Servers) {
            $ServerSGPropertyArray = Get-StringPropertySplit -SearchString ($Server -Replace 'add server' ,'') -RemoveQuotes;    
            $SRVSGNAME = $ServerSGPropertyArray[0];
            
            foreach ($SVCGroupBind in $ServiceGroupBindMatches) {
                $SGBINDNAME = Get-StringPropertySplit -SearchString ($SVCGroupBind -Replace 'bind serviceGroup' ,'') -RemoveQuotes;    
                $SGBINDNAME = $SGBINDNAME[1]

                If ($SRVSGNAME -eq $SGBINDNAME) {
                    $ServiceGroupServers += @{ Server = $SRVSGNAME; }
                }
            }
        }
        $ServiceGroupServers
        if ($ServiceGroupServers.Length -gt 0) {
            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ServiceGroupServers;                   
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
            FindWordDocumentEnd;
        } else {
            WriteWordLine 0 0 "No Server has been configured for this Service Group"
        }   

        WriteWordLine 0 0 " "

        WriteWordLine 3 0 "Monitor"     

        if ($Null -eq $ServicegroupBindMatches -or $ServicegroupBindMatches.Length -le 0) {
            WriteWordLine 0 0 "No Monitor has been configured for this Service Group"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ServiceGroupMonitors = @();

            ## IB - Iterate over all Service bindings (uses new function)
            foreach ($SVCGroupBind in $ServiceGroupBindMatches) {

                if (Test-StringProperty -SearchString $SVCGroupBind -PropertyName "-monitorName") {
                    $ServiceGroupMonitors += @{ Monitor = Get-StringProperty $SVCGroupBind "-monitorName"; }
                }
            } # end foreach

            if ($ServiceGroupMonitors.Length -gt 0) {
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $ServiceGroupMonitors;                   
                    AutoFit = $wdAutoFitContent;
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
                
                FindWordDocumentEnd;
            } else {
                WriteWordLine 0 0 "No Monitor has been configured for this Service Group"
        }   
        } # end if

        WriteWordLine 0 0 " "
        WriteWordLine 3 0 "Advanced Configuration"

        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $AdvancedConfiguration = @(
            @{ Description = "Description"; Value = "Configuration"; }
            @{ Description = "Clear text port"; Value = Get-StringProperty $ServiceGroup "-clearTextPort" "NA" ; }
			@{ Description = "Cache Type"; Value = Get-StringProperty $ServiceGroup "-cacheType" "NA" ; }
			@{ Description = "Maximum Client Requests"; Value = Get-StringProperty $ServiceGroup "-maxClient" "4294967294 (Maximum Value)" ; }
			@{ Description = "Monitor health of this Service Group"; Value = Test-NotStringPropertyYesNo $ServiceGroup "-healthMonitor" ; }
			@{ Description = "Maximum Requests"; Value = Get-StringProperty $ServiceGroup "-maxreq" "65535 (Maximum Value)" ; }
			@{ Description = "Use Transparent Cache"; Value = Test-StringPropertyYesNo $ServiceGroup "-cacheable" ; }
			@{ Description = "Insert the Client IP header"; Value = Get-StringProperty $ServiceGroup "-cip" "NA"  ; }
			@{ Description = "Name for the HTTP header"; Value = Get-StringProperty $ServiceGroup "-cipHeader" "NA" ; }
			@{ Description = "Use Source IP"; Value = Test-StringPropertyYesNo $ServiceGroup "-usip" ; }
            @{ Description = "Path Monitoring"; Value = Test-StringPropertyYesNo $ServiceGroup "-pathMonitor" ; }
			@{ Description = "Individual Path monitoring"; Value = Test-StringPropertyYesNo $ServiceGroup "-pathMonitorIndv" ; }
			@{ Description = "Use the proxy port"; Value = Test-StringPropertyYesNo $ServiceGroup "-useproxyport" ; }
			@{ Description = "SureConnect"; Value = Test-StringPropertyOnOff $ServiceGroup "-sc" ; }
			@{ Description = "Surge protection"; Value = Test-StringPropertyOnOff $ServiceGroup "-sp" ; }
			@{ Description = "RTSP session ID mapping"; Value = Test-StringPropertyOnOff $ServiceGroup "-rtspSessionidRemap" ; }
			@{ Description = "Client Time-Out"; Value = Get-StringProperty $ServiceGroup "-cltTimeout" "31536000 (Maximum Value)" ; }
			@{ Description = "Server Time-Out"; Value = Get-StringProperty $ServiceGroup "-svrTimeout" "3153600 (Maximum Value)" ; }
			@{ Description = "Unique identifier for the Service Group"; Value = Get-StringProperty $ServiceGroup "-CustomServerID" "None" ; }
			@{ Description = "The identifier for the Service Group"; Value = Get-StringProperty $ServiceGroup "-serverID" "None" ; }
			@{ Description = "Enable client keep-alive"; Value = Test-StringPropertyYesNo $ServiceGroup "-CKA" ; }
			@{ Description = "Enable TCP buffering"; Value = Test-StringPropertyYesNo $ServiceGroup "-TCPB" ; }
            @{ Description = "Enable compression"; Value = Test-StringPropertyYesNo $ServiceGroup "-CMP" ; }
			@{ Description = "Maximum bandwidth, in Kbps"; Value = Get-StringProperty $ServiceGroup "-maxBandwidth" "4294967287 (Maximum Value)" ; }
			@{ Description = "Use Layer 2 mode"; Value = Test-StringPropertyYesNo $ServiceGroup "-accessDown" ; }
			@{ Description = "Sum of weights of the monitors"; Value = Get-StringProperty $ServiceGroup "-monThreshold" "65535 (Maximum Value)" ; }
			@{ Description = "Initial state of the Service Group"; Value = Test-NotStringPropertyEnabledDisabled $ServiceGroup "-state" ; }
			@{ Description = "Perform delayed clean-up"; Value = Test-NotStringPropertyEnabledDisabled $ServiceGroup "-downStateFlush" ; }
			@{ Description = "TCP profile"; Value = Get-StringProperty $ServiceGroup "-tcppProfileName" "NA" ; }
			@{ Description = "HTTP profile"; Value = Get-StringProperty $ServiceGroup "-httpProfileName" "NA" ; }
			@{ Description = "A numerical identifier"; Value = Get-StringProperty $ServiceGroup "-hashId" "NA" ; }
			@{ Description = "Comment about the ServiceGroup"; Value = Get-StringProperty $ServiceGroup "-comment" "NA"; }
			@{ Description = "Logging of AppFlow information"; Value = Test-NotStringPropertyEnabledDisabled $ServiceGroup "-appflowLog" ; }
			@{ Description = "Network profile"; Value = Get-StringProperty $ServiceGroup "-netProfile" "NA" ; }
        );

        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = $AdvancedConfiguration;
            Columns = "Description","Value";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines -List;

		FindWordDocumentEnd;
		$TableRange = $Null
		$Table = $Null
        WriteWordLine 0 0 " "

        $selection.InsertNewPage() 
        }
   }
$selection.InsertNewPage() 
#endregion NetScaler Service Groups

#region NetScaler Servers
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Servers"
WriteWordLine 1 0 "NetScaler Servers"

if($Servers.Length -le 0) { WriteWordLine 0 0 "No Server has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $ServersH = @();

    foreach ($Server in $Servers) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $ServerPropertyArray = Get-StringPropertySplit -SearchString ($Server -Replace 'add Server' ,'') -RemoveQuotes;

        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $ServersH += @{
                Server = $ServerPropertyArray[0];
                IP = $ServerPropertyArray[1];
                TD = Get-StringProperty $Server "-td" "0 (Default)";
                STATE = Test-NotStringPropertyEnabledDisabled $Server "-state";
                COMMENT = Get-StringProperty $Server "-comment" "No Comment" -RemoveQuotes;
            }
        }
        if ($ServersH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $ServersH;
                Columns = "Server","IP","TD","STATE","COMMENT";
                Headers = "Server","IP Address","Traffic Domain","State","Comment";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

$selection.InsertNewPage()    
#endregion NetScaler Servers

#endregion traffic management

#region Citrix NetScaler Gateway

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway"
WriteWordLine 1 0 "Citrix NetScaler (Access) Gateway"

#region Citrix NetScaler Gateway CAG Global...

WriteWordLine 2 0 "NetScaler Gateway Global Settings"
Write-Verbose "$(Get-Date): `tNetScaler Gateway Global Settings"
#region GlobalNetwork

WriteWordLine 3 0 "Global Settings Network"

## IB - Create an array of hashtables to store our columns. Note: If we need the
## IB - headers to include spaces we can override these at table creation time.
## IB - Create the parameters to pass to the AddWordTable function

ForEach ($LINE in $SetVpnParameter) {
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            Wins = Get-StringProperty $LINE "-winsIP" "Not Configured";
            Mapped = Get-StringProperty $LINE "-useMIP" "VPN_SESS_ACT_NS";
            Intranet = Get-StringProperty $LINE "-iipDnsSuffix" "Not Configured";
            Http = Get-StringProperty $LINE "-httpPort" "Not Configured";
            Timeout = Get-StringProperty $LINE "-forcedTimeout" "Not Configured";
        }
        Columns = "Wins","Mapped","Intranet","Http","Timeout";
        Headers = "WINS Server","Mapped IP","Intranet IP","HTTP Ports","Forced Time-out";
        Format = -235; ## IB - Word constant for Light List Accent 5
        AutoFit = $wdAutoFitContent;
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    }

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#endregion GlobalNetwork

#region GlobalClientExperience
WriteWordLine 3 0 "Global Settings Client Experience"

ForEach ($LINE in $SetVpnParameter) {

    ## IB - Create an array of hashtables to store our columns.
    ## IB - about column names as we'll utilise a -List(view)!
    [System.Collections.Hashtable[]] $NsGlobalClientExperience = @(
        ## IB - Each hashtable is a separate row in the table!
        @{ Column1 = "Description"; Column2 = "Value"; }
        @{ Column1 = "Home Page"; Column2 = Get-StringProperty $LINE "-homePage" "Not Configured"; }
        @{ Column1 = "URL for Web Based E-mail"; Column2 = Get-StringProperty $LINE "-useMIP" "Not Configured"; }
        @{ Column1 = "Split Tunnel"; Column2 = Get-StringProperty $LINE "-splitTunnel" "Off"; }
        @{ Column1 = "Session Time-Out"; Column2 = Get-StringProperty $LINE "-sessTimeout" "0"; }
        @{ Column1 = "Client-Idle Time-Out"; Column2 = Get-StringProperty $LINE "-clientIdleTimeout" "0"; }
        @{ Column1 = "Plug-in Type"; Column2 = Get-StringProperty $LINE "-epaClientType" "AGENT"; }
        @{ Column1 = "Clientless Access"; Column2 = Get-StringProperty $LINE "-clientlessVpnMode" "Off"; }
        @{ Column1 = "Clientless URL Encoding"; Column2 = Get-StringProperty $LINE "-clientlessModeUrlEncoding" "VPN_SESS_ACT_CVPN_ENC_OPAQUE"; }
        @{ Column1 = "Clientless Persistent Cookie"; Column2 = Get-StringProperty $LINE "-clientlessPersistentCookie" "Deny"; }
        @{ Column1 = "Single Sign-On to Web Applications"; Column2 = Get-StringProperty $LINE "-SSO" "Off"; }
        @{ Column1 = "Credential Index"; Column2 = Get-StringProperty $LINE "-ssoCredential" "Primary"; }
        @{ Column1 = "KCD Account"; Column2 = Get-StringProperty $LINE "-kcdAccount" "Not Configured"; }
        @{ Column1 = "Single Sign-On with Windows"; Column2 = Get-StringProperty $LINE "-windowsAutoLogon" "Off"; }
        @{ Column1 = "Client Cleanup Prompt"; Column2 = Get-StringProperty $LINE "-forceCleanup" "Off"; }
        @{ Column1 = "UI Theme"; Column2 = Get-StringProperty $LINE "-UITHEME" "DEFAULT"; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $NsGlobalClientExperience;
        Columns = "Column1","Column2";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    $Table = AddWordTable @Params -List -NoGridLines;
    }
FindWordDocumentEnd;

$NsGlobalClientExperience = $null;

WriteWordLine 0 0 " "
#endregion GlobalClientExperience

#region GlobalSecurity
WriteWordLine 3 0 "Global Settings Security"

ForEach ($LINE in $SetVpnParameter) {

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            DEFAUTH = Get-StringProperty $LINE "-defaultAuthorizationAction" "DENY";
            CLISEC = Get-StringProperty $LINE "-encryptCsecExp" "Disabled";
            SECBRW = Get-StringProperty $LINE "-SecureBrowse" "Enabled";
        }
        Columns = "DEFAUTH","CLISEC","SECBRW";
        Headers = "Default Authorization Action","Client Security Encryption","Secure Browse";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

WriteWordLine 0 0 " "
#endregion GlobalSecurity

#region GlobalPublishedApps
WriteWordLine 3 0 "Global Settings Published Applications"

ForEach ($LINE in $SetVpnParameter) {

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ICAPROXY = Get-StringProperty $LINE "-icaProxy" "OFF";
            WIADDR = Get-StringProperty $LINE "-wihome" "Not Configured" -RemoveQuotes;
            WIMODE = Get-StringProperty $LINE "-wiPortalMode" "NORMAL";
            SSO = Get-StringProperty $LINE "-ntDomain" "Not Configured";
            HOME = Get-StringProperty $LINE "-citrixReceiverHome" "Not Configured";
            ACCSVC = Get-StringProperty $LINE "-storefronturl" "Not Configured";
        }
        Columns = "ICAPROXY","WIADDR","WIMODE","SSO","HOME","ACCSVC";
        Headers = "ICA Proxy","Web Interface addres","Web Interface Portal Mode","Single Sign-On Domain","Citrix Receiver Home Page","Account Services Address";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

WriteWordLine 0 0 " "

#endregion GlobalPublishedApps

#region Global PREAUTH
WriteWordLine 3 0 "Global Settings Pre-Authentication Settings"
if($SETAAAPREAUTH.Length -le 0) { $SETAAAPREAUTH = "123"}

ForEach ($AAAAUTH in $SETAAAPREAUTH) {
    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - This table will only have 1 row so create the nested hashtable inline
            ACTION = Get-StringProperty $AAAAUTH "-preauthenticationaction" "ALLOW";
            PROC1 = Get-StringProperty $AAAAUTH "-killProcess" "Not Configured";
            FILES1 = Get-StringProperty $AAAAUTH "-deletefiles" "Not Configured";
            Expr1 = Get-StringProperty $AAAAUTH "-rule" "Not Configured";
        }
        Columns = "ACTION","PROC1","FILES1","Expr1";
        Headers = "Action","Processes to be cancelled","Files to be deleted","Expression";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
    }
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    }

WriteWordLine 0 0 " "
#endregion Global PREAUTH

#region GlobalAuthentication
WriteWordLine 3 0 "Global Settings Authentication Settings"

$Set | ForEach-Object {  
    if ($_ -like 'set aaa parameter *') {
        ## IB - Create an array of hashtables to store our columns. Note: If we need the
        ## IB - headers to include spaces we can override these at table creation time.
        ## IB - Create the parameters to pass to the AddWordTable function
        $Params = $null
        $Params = @{
            Hashtable = @{
                ## IB - Each hashtable is a separate row in the table!
                MAXUSR = Get-StringProperty $_ "-maxAAAUsers" "1";
                NATIP = Get-StringProperty $_ "-aaadnatIp" "Default Setting";
                MAXLOG = Get-StringProperty $_ "-maxLoginAttempts" "Unlimited";
                FAILTO = Get-StringProperty $_ "-failedLoginTimeout" "Default Setting";
                ENSTAT = Get-StringProperty $_ "-enableStaticPageCaching" "Enabled";
                ENADV = Get-StringProperty $_ "-enableEnhancedAuthFeedback" "Disabled";
                DEFAUTH = Get-StringProperty $_ "-defaultAuthType" "Local Authentication";
            }
            Columns = "MAXUSR","NATIP","MAXLOG","FAILTO","ENSTAT","ENADV","DEFAUTH";
            Headers = "Maximum Number of Users","NAT IP Address","Maximum login Attempts","Failed Login Timeout","Enable Static Caching","Enable advanced authentication feedback","Default Authentication Type";
            AutoFit = $wdAutoFitContent;
            Format = -235; ## IB - Word constant for Light List Accent 5
        }
        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        }
    }

FindWordDocumentEnd;

WriteWordLine 0 0 " "
#endregion GlobalAuthentication

#region Global STA
WriteWordLine 3 0 "Global Settings Secure Ticket Authority Configuration"
$STAMATCHES = Get-StringWithProperty -SearchString $Bind -Like "bind vpn global -staServer *";

if ($Null -eq $STAMATCHES -or $STAMatches.Length -le 0) {
            WriteWordLine 0 0 "No Secure Ticket Authority has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $STAS = @();

                ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($STALINE in $STAMATCHES) {
                $STAS += @{ STA = Get-StringProperty $STALINE "-staServer" -RemoveQuotes; }
                } # end foreach
            $Params = $null
            $Params = @{
                Hashtable = $STAS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
FindWordDocumentEnd;
WriteWordLine 0 0 " "

#endregion Global STA

#region Global AppController
WriteWordLine 3 0 "Global Settings App Controller Configuration"
$APPCMATCHES = Get-StringWithProperty -SearchString $Bind -Like "bind vpn global -appController *";

if ($Null -eq $APPCMATCHES -or $APPCMatches.Length -le 0) {
            WriteWordLine 0 0 "No App Controller has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $APPCS = @();

                ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($APPCLINE in $APPCMATCHES) {
                $APPCS += @{ "APP Controller" = Get-StringProperty $APPCLINE "-appController" -RemoveQuotes; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $APPCS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
#endregion Global AppController

$selection.InsertNewPage()

#endregion CAG Global

#region CAG vServers

if($AccessGateways.Length -le 0) { WriteWordLine 0 0 "No Citrix NetScaler Gateway has been configured"} else {
    $CurrentRowIndex = 0;
    foreach ($AccessGateway in $AccessGateways) {
        $CurrentRowIndex++;

        ## IB - Store the Load Balancer as native and escaped. This is needed by LBBIND etc.
        $vServerDisplayName = Get-StringProperty $AccessGateway "vserver";
        $vServerDisplayNameNoQuotes = Get-StringProperty $AccessGateway "vserver" -RemoveQuotes;
        $vServerName = $vServerDisplayName.Trim();

        WriteWordLine 2 0 "NetScaler Gateway Virtual Server: $(Get-StringProperty $AccessGateway "vserver" -RemoveQuotes)";
        Write-Verbose "$(Get-Date): `tNetScaler Gateway $CurrentRowIndex/$($AccessGateways.Length) : $vServerDisplayNameNoQuotes";

#region CAG vServer basic configuration

        $AGPropertyArray = Get-StringPropertySplit -SearchString ($AccessGateway -Replace 'add vpn vserver' ,'') -RemoveQuotes;

        ## IB - Create an array of hashtables to store our columns. Note: If we need the
        $Params = $null
        $Params = @{
            Hashtable = @{
                State = Get-StringProperty $AccessGateway "-state" "Enabled";
                Mode = Test-NotStringPropertyOnOff $AccessGateway "-icaOnly";
                IPAddress = $AGPropertyArray[2];
                Port = $AGPropertyArray[3];
                Protocol = $AGPropertyArray[1];
                MaximumUsers = Get-StringProperty $AccessGateway "-maxAAAUsers" "Unlimited";
                MaxLogin = Get-StringProperty $AccessGateway "-maxLoginAttempts" "Unlimited";
            }
            Columns = "State","Mode","IPAddress","Port","Protocol","MaximumUsers","MaxLogin";
            Headers = "State","Smart Access","IP Address","Port","Protocol","Maximum Users","Maximum Logons";
            AutoFit = $wdAutoFitContent
            Format = -235; ## IB - Word constant for Light List Accent 5
        }

        ## IB - Add the table to the document, splatting the parameters
        $Table = AddWordTable @Params -NoGridLines;
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "
#endregion CAG vServer basic configuration

#region CAG certificate
        
        WriteWordLine 3 0 "Certificates"
        $CAGSERVERCERTBINDS = Get-StringWithProperty -SearchString $CERTBINDS -Like "bind ssl vserver $vServerDisplayName -certkeyName *";
        if($CAGSERVERCERTBINDS.Length -le 0) { WriteWordLine 0 0 "No Certificate has been configured for this NetScaler Gateway vServer"} else { 
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $CAGCERTSH = @();
            
            Foreach ($CERT in $CAGSERVERCERTBINDS) {$CAGCERTSH += @{ Certificate = Get-StringProperty $CERT "-certkeyName" -RemoveQuotes; }}

            if ($CAGCERTSH.Length -gt 0) {
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $CAGCERTSH;
                    AutoFit = $wdAutoFitContent;
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " " 
            }
        }
#endregion CAG certificate

#region CAG vServer policies
        $CAGPOLS = Get-StringWithProperty -SearchString $BINDVPNVSERVER -Like "bind vpn vserver $vServerDisplayName *"; 
    
    #region CAG Authentication LDAP Policies        
        
        WriteWordLine 3 0 "Authentication LDAP Policies"
        
        if($AUTHLDAPPOLS.Length -le 0) { WriteWordLine 0 0 "No LDAP Policy has been configured"} else { 

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLHASH = @();

            ## BS 1. First we get all LDAP Policies on the NetScaler system
            foreach ($AUTHLDAPPOL in $AUTHLDAPPOLS) {
                $CAGAUTHPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHLDAPPOL -Replace 'add authentication ldapPolicy' ,'') ;
                $CAGAUTHPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($AUTHLDAPPOL -Replace 'add authentication ldapPolicy' ,'') -RemoveQuotes ;
                
                ## BS 2. Then we determine the policy name for each of the LDAP Policies
                $POLICYNAME = $CAGAUTHPOLPropertyArray[0];
                $LDAPPolicyDisplayName = Get-StringProperty $AUTHLDAPPOL "ldapPolicy" -RemoveQuotes;

                ## BS 3. Now we find out if this specific LDAP policy is bound to this specific CAG vServer
                $AUTHPOLS = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -policy $POLICYNAME*";

                ## BS 4. If we have an LDAP Policy bound to this specific CAG vServer then we 
                foreach ($AUTHPOL in $AUTHPOLS) {                
                    $PRIMARY = Test-StringProperty $AUTHPOL -PropertyName "-secondary";
                    If ($PRIMARY -eq $True) {$PRIMARY = "Secondary"} else {$PRIMARY = "Primary"}

                    $AUTHPOLHASH += @{
                        Name = $LDAPPolicyDisplayName;
                        Action = $CAGAUTHPOLPropertyArrayNoQuotes[2];
                        Expr = $CAGAUTHPOLPropertyArrayNoQuotes[1];
                        Primary = $PRIMARY ;
                        Priority = Get-StringProperty $AUTHPOL "-priority";
                    } # end Hasthable $AUTHPOLH1
                }# end foreach $AUTHPOLS
            } #end foreach AUTHLDAPPOLS

            if ($AUTHPOLHASH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $AUTHPOLHASH;
                    Columns = "Name","Action","Expr","Primary","Priority";
                    Headers = "Policy Name","Policy Action","Expression","Primary","Priority";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;
                ## IB - Set the header background and bold font
                #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;

                FindWordDocumentEnd;
                
            } else { WriteWordLine 0 0 "No LDAP Policy has been configured"} #endif AUTHPOLHASH.Length
        } #end if no LDAP configures
    WriteWordLine 0 0 " "
    #endregion CAG Authentication LDAP Policies  

    #region CAG Authentication Radius Policies        
        
        WriteWordLine 3 0 "Authentication Radius Policies"
        if($AUTHRADPOLS.Length -le 0) { WriteWordLine 0 0 "No Radius Policy has been configured"} else {        
            $AUTHPOLRADHASH = $null
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $AUTHPOLRADHASH = @();

            ## BS 1. First we get all RADIUS Policies on the NetScaler system
            foreach ($AUTHRADPOL in $AUTHRADPOLS) {
                $CAGAUTHRADPOLPropertyArray = Get-StringPropertySplit -SearchString ($AUTHRADPOL -Replace 'add authentication radiusPolicy' ,'') ;
                $CAGAUTHRADPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($AUTHRADPOL -Replace 'add authentication radiusPolicy' ,'') -RemoveQuotes ;
                
                ## BS 2. Then we determine the policy name for each of the LDAP Policies
                $POLICYNAME = $null
                $POLICYNAME = $CAGAUTHRADPOLPropertyArray[0];

                ## BS 3. Now we find out if this specific RADIUS policy is bound to this specific CAG vServer
                $AUTHRADPOLSBIND = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -policy $POLICYNAME*";

                ## BS 4. If we have an RADIUS Policy bound to this specific CAG vServer then we 
                $AUTHPOL = $null
                foreach ($AUTHPOL in $AUTHRADPOLSBIND) {                
                    $PRIMARY = Test-StringProperty $AUTHPOL -PropertyName "-secondary";
                    If ($PRIMARY -eq $True) {$PRIMARY = "Secondary"} else {$PRIMARY = "Primary"}
                    
                    $AUTHPOLRADHASH += @{
                        Name = $CAGAUTHRADPOLPropertyArrayNoQuotes[0];
                        Action = $CAGAUTHRADPOLPropertyArrayNoQuotes[2];
                        Expr = $CAGAUTHRADPOLPropertyArrayNoQuotes[1];
                        Primary = $PRIMARY ;
                        Priority = Get-StringProperty $AUTHPOL "-priority";
                    } # end Hasthable $AUTHPOLRADHASH
                }# end foreach $AUTHPOLS
            } #end foreach AUTHRADPOLS

            if ($AUTHPOLRADHASH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $AUTHPOLRADHASH;
                    Columns = "Name","Action","Expr","Primary","Priority";
                    Headers = "Policy Name","Policy Action","Expression","Primary","Priority";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No Radius Policy has been configured"} #endif AUTHPOLHASH.Length
        } #End If no policies configured
    WriteWordLine 0 0 " "
    #endregion CAG Authentication Radius Policies  
    
    #region CAG Session Policies        
       
        WriteWordLine 3 0 "Session Policies"
        if($CAGSESSIONPOLS.Length -le 0) { WriteWordLine 0 0 "No Session Policy has been configured"} else { 

            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $SESSIONPOLH = @();

            ## BS 1. First we get all Session Policies on the NetScaler system
            foreach ($CAGSESSIONPOL in $CAGSESSIONPOLS) {
                $CAGSESSIONPOLPropertyArray = $Null
                $CAGSESSIONPOLPropertyArray = Get-StringPropertySplit -SearchString ($CAGSESSIONPOL -Replace 'add vpn' ,'') ;
                $CAGSESSIONPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($CAGSESSIONPOL -Replace 'add vpn' ,'') -RemoveQuotes ;
                ## BS 2. Then we determine the policy name for each of the Session Policies
                $POLICYNAME = $CAGSESSIONPOLPropertyArray[1];
                
                ## BS 3. Now we find out if this specific Session policy is bound to this specific CAG vServer
                $SESSIONPOLS = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -policy $POLICYNAME*";

                ## BS 4. If we have an Session Policy bound to this specific CAG vServer then we 
                foreach ($SESSIONPOL in $SESSIONPOLS) {                
                    $SESSIONPOLH += @{
                        Name = Get-StringProperty $SESSIONPOL "-policy" -RemoveQuotes;
                    } # end Hasthable $SESSIONPOLH
                }# end foreach $SESSIONPOLS
            } #end foreach SESSIONPOLS
            
            if ($SESSIONPOLH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $SESSIONPOLH;
                        Columns = "Name";
                        Headers = "Policy Name";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No Session Policy has been configured"} #endif SESSIONPOLHASH.Length
        } #end if no Session policy configures
WriteWordLine 0 0 " "
    #endregion CAG Session Policies 
        
    #region CAG URL Bookmarks   
       
        WriteWordLine 3 0 "URL Bookmarks "

        if($CAGURLPOLS.Length -le 0) { WriteWordLine 0 0 "No URL Bookmark has been configured"} else { 
            $URLPOLH = $null
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $URLPOLH = @();

            ## BS 1. First we get all URL  Policies on the NetScaler system
            foreach ($CAGURLPOL in $CAGURLPOLS) {
                $CAGURLPOLPropertyArray = Get-StringPropertySplit -SearchString ($CAGURLPOL -Replace 'add vpn' ,'') ;
                $CAGURLPOLPropertyArrayNoQuotes = Get-StringPropertySplit -SearchString ($CAGURLPOL -Replace 'add vpn' ,'') -RemoveQuotes ;
                ## BS 2. Then we determine the policy name for each of the URL Policies
                $POLICYNAME = $CAGURLPOLPropertyArray[1];
                
                ## BS 3. Now we find out if this specific URL policy is bound to this specific CAG vServer
                $URLPOLS = Get-StringWithProperty -SearchString $CAGPOLS -Like "bind vpn vserver $vServerDisplayName -urlName $POLICYNAME*";

                $TEXTTODISPLAY = $CAGURLPOLPropertyArrayNoQuotes[2]
                $URL = $CAGURLPOLPropertyArrayNoQuotes[3]

                ## BS 4. If we have an URL Policy bound to this specific CAG vServer then we 
                foreach ($URLPOL in $URLPOLS) {                
                    $URLPOLH += @{
                        Name = Get-StringProperty $URLPOL "-urlName" -RemoveQuotes;
                        Text = $TEXTTODISPLAY;
                        URL = $URL;
                        CLIENTLESSACCESS = Get-StringProperty $CAGURLPOL "-clientlessAccess" "Off" -RemoveQuotes;
                        Comment = Get-StringProperty $CAGURLPOL "-comment" "No Comment" -RemoveQuotes;
                    } # end Hasthable $URLPOLH
                }# end foreach $SESSIONPOLS
            } #end foreach SESSIONPOLS
            if ($URLPOLH.Length -gt 0) {
                ## IB - Add the table to the document (only if not null!
                ## IB - Create the parameters to pass to the AddWordTable function
                $Params = $null
                $Params = @{
                    Hashtable = $URLPOLH;
                        Columns = "Name","Text","URL","CLIENTLESSACCESS","Comment";
                        Headers = "Policy Name","Text to display","URL","Clientless Access","Comment";
                    AutoFit = $wdAutoFitContent
                    Format = -235; ## IB - Word constant for Light List Accent 5
                }
                ## IB - Add the table to the document, splatting the parameters
                $Table = AddWordTable @Params -NoGridLines;

                FindWordDocumentEnd;
            } else { WriteWordLine 0 0 "No URL has been configured"} #endif SESSIONPOLHASH.Length
        } #end if no Session policy configures
    WriteWordLine 0 0 " "
    #endregion CAG URL Policies 

    #region CAG STA Configuration

    WriteWordLine 3 0 "Secure Ticket Authority Configuration"
    $vServerSTAs = Get-StringWithProperty -SearchString $BINDVPNVSERVER -Like "bind vpn vserver $vServerDisplayName -staServer*";    
    if($vServerSTAs.Length -gt 0) {
        $vServerSTAH = $null
        
        ## IB - Use an array of hashtable to store the rows
        [System.Collections.Hashtable[]] $vServerSTAH = @();
 
        foreach ($vServerSTA in $vServerSTAs) {
        $vServerSTAH += @{
            Name = Get-StringProperty $vServerSTA "-staServer" -RemoveQuotes;
            } 
        }

        if ($vServerSTAH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $vServerSTAH;
                    Columns = "Name";
                    Headers = "Security Ticket Authority";
                AutoFit = $wdAutoFitContent
                Format = -235; ## IB - Word constant for Light List Accent 5
            }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
        } else {
            WriteWordLine 0 0 "No specific Secure Ticket Authority has been configured for this virtual server"
        }
    WriteWordLine 0 0 " "
    } # if($vServerSTAs.Length
    #endregion CAG STA Configuration
#endregion CAG vServer policies
          
        $selection.InsertNewPage()
    } #end foreach AccessGateway
} #end if Accessgateway.Length

#endregion CAG vServers

#region CAG Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters Citrix NetScaler (Access) Gateway Policies"
WriteWordLine 1 0 "NetScaler Gateway Policies"

#region CAG Session Policies
WriteWordLine 0 0 " "
WriteWordLine 2 0 "NetScaler Gateway Session Policies"
WriteWordLine 0 0 " "
Write-Verbose "$(Get-Date): `tNetScaler Gateway Session Policies"

$CurrentRowIndex = 0;

ForEach ($CAGSESSIONACT in $CAGSESSIONACTS) {
    $CurrentRowIndex++
    WriteWordLine 3 0 "NetScaler Gateway Session Policy: $(Get-StringProperty $CAGSESSIONACT "sessionAction" -RemoveQuotes)";
    Write-Verbose "$(Get-Date): `t`tNetScaler Gateway $CurrentRowIndex/$($CAGSESSIONACTS.Length) : $(Get-StringProperty $CAGSESSIONACT "sessionAction" -RemoveQuotes)";
#region Security
    
    WriteWordLine 4 0 "Security"

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            DEFAUTH = Get-StringProperty $CAGSESSIONACT "-defaultAuthorizationAction" "DENY";
            CLISEC = Get-StringProperty $CAGSESSIONACT "-encryptCsecExp" "Disabled";
            SECBRW = Get-StringProperty $CAGSESSIONACT "-SecureBrowse" "Disabled";
        }
        Columns = "DEFAUTH","CLISEC","SECBRW";
        Headers = "Default Authorization Action","Client Security Encryption","Secure Browse";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
#endregion Security

#region Published Applications  

    WriteWordLine 4 0 "Published Applications"

    ## IB - Create an array of hashtables to store our columns. Note: If we need the
    ## IB - headers to include spaces we can override these at table creation time.
    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = @{
            ## IB - Each hashtable is a separate row in the table!
            ICAPROXY = Get-StringProperty $CAGSESSIONACT "-icaProxy" "Global Configuration";
            WIADDR = Get-StringProperty $CAGSESSIONACT "-wihome" "Global Configuration" -RemoveQuotes;
            WIMODE = Get-StringProperty $CAGSESSIONACT "-wiPortalMode" "Global Configuration";
            SSO = Get-StringProperty $CAGSESSIONACT "-ntDomain" "Global Configuration";
            HOME = Get-StringProperty $CAGSESSIONACT "-citrixReceiverHome" "Global Configuration";
            ACCSVC = Get-StringProperty $CAGSESSIONACT "-storefronturl" "Global Configuration";
        }
        Columns = "ICAPROXY","WIADDR","WIMODE","SSO","HOME","ACCSVC";
        Headers = "ICA Proxy","Web Interface addres","Web Interface Portal Mode","Single Sign-On Domain","Citrix Receiver Home Page","Account Services Address";
        AutoFit = $wdAutoFitContent;
        Format = -235; ## IB - Word constant for Light List Accent 5
    }

    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -NoGridLines;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "

#end region Published Applications
    $selection.InsertNewPage()
}

    #endregion CAG Session Policies

#endregion CAG Session Policies

#endregion CAG Policies

#endregion Citrix NetScaler Gateway

#region NetScaler Monitors
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Monitors"

WriteWordLine 1 0 "NetScaler Custom Monitors"

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler Monitors Table"

if($MONITORS.Length -le 0) { WriteWordLine 0 0 "No Custom Monitor has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $MONITORSH = @();

    foreach ($MONITOR in $MONITORS) {

        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $MONITORPropertyArray = Get-StringPropertySplit -SearchString ($MONITOR -Replace 'add lb monitor ' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $MONITORSH += @{
                NAME = $MONITORPropertyArray[0];
                Protocol = $MONITORPropertyArray[1];
                HTTPRequest = Get-StringProperty $MONITOR "-httpRequest" "NA";
                DestinationIP = Get-StringProperty $MONITOR "-destIP" "NA";
                DestinationPort = Get-StringProperty $MONITOR "-destPort" "NA";
                Interval = Get-StringProperty $MONITOR "-interval" "NA";
                ResponseCode = Get-StringProperty $MONITOR "-respCode" "NA";
                TimeOut = Get-StringProperty $MONITOR "-resptimeout" "NA";
                SitePath = Get-StringProperty $MONITOR "-sitePath" "NA";
                }
            }

        if ($MONITORSH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $MONITORSH;
                Columns = "NAME","Protocol","HTTPRequest","DestinationIP","DestinationPort","Interval","ResponseCode","TimeOut","SitePath";
                Headers = "Monitor Name","Protocol","HTTP Request","Destination IP","Destination Port","Interval","Response Code","Time-Out","SitePath";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }

$selection.InsertNewPage()

#endregion NetScaler Monitors

#region NetScaler Policies
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Policies"

WriteWordLine 1 0 "NetScaler Policies"

## Work in Progress: Binding to actions and binding to vServers

#Policy Pattern Set
WriteWordLine 2 0 "NetScaler Custom Pattern Set Policies"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Pattern Set Policies"
$PATSET1 = Get-StringWithProperty -SearchString $Add -Like 'add policy patset *';

if ($Nul -eq $PATSET1 -or $PATSET1.Length -le 0) {
        WriteWordLine 0 0 "No Custom Pattern Set Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $PATSETS = @();

            foreach ($PAT in $PATSET1) {
                $Y = ($PAT -replace 'add policy patset ', '').split()
                $PATSETS += @{ "Pattern Set Policy" = "$Y"; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $PATSETS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Custom Responder Policies"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Responder Policies"
$POLICY = Get-StringWithProperty -SearchString $Add -Like 'add responder policy *';

if ($Null -eq $POLICY -or $POLICY.Length -le 0) {
        WriteWordLine 0 0 "No Custom Responder Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $POLICIESH = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($POL in $POLICY) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes
                $POLICIESH += @{ "Responder Policy" = $Y[3]; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $PARAMS = $null
            $Params = @{
                Hashtable = $POLICIESH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Custom Rewrite Policies"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Rewrite Policies"

$POLRW = Get-StringWithProperty -SearchString $Add -Like 'add rewrite policylabel *';

if ($Null -eq $POLRW -or $POLRW.Length -le 0) {
        WriteWordLine 0 0 "No Custom Rewrite Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $POLRWH = @();

            foreach ($POL in $POLRW) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes
                $POLRWH += @{ "Rewrite Policy" = $Y[3]; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $POLRWH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            ## IB - Set the header background and bold font
            #SetWordCellFormat -Collection $Table.Rows.First.Cells -BackgroundColor $wdColorGray15 -Bold;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion NetScaler Policies

#region NetScaler Actions
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Actions"

WriteWordLine 1 0 "NetScaler Actions"

## Work in Progress: Binding to policies

WriteWordLine 2 0 "NetScaler Custom Pattern Set Action"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Pattern Set Action"
$ACTPATSET1 = Get-StringWithProperty -SearchString $Add -Like 'add action patset *';

if ($Null -eq $ACTPATSET1 -or $ACTPATSET1.Length -le 0) {
        WriteWordLine 0 0 "No Custom Pattern Set Action has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ACTPATSETS = @();

            foreach ($ACTPAT in $ACTPATSET1) {
                $Y = ($ACTPAT -replace 'add action patset ', '').split()
                $ACTPATSETS += @{ "Pattern Set Policy" = "$Y"; }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ACTPATSETS;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Responder Action"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Responder Action"
$ACTRES = Get-StringWithProperty -SearchString $Add -Like 'add responder action *';

if ($Null -eq $ACTRES -or $ACTRES.Length -le 0) {
        WriteWordLine 0 0 "No Custom Responder Action has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ACTRESH = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($POL in $ACTRES) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes                
                $ACTRESH += @{ 
                    Responder = $Y[3]; 
                    Rule = $Y[4];
                    Undefined = $Y[5];
                    }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ACTRESH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

WriteWordLine 2 0 "NetScaler Custom Rewrite Action"

Write-Verbose "$(Get-Date): `tTable: NetScaler Custom Rewrite Action"
$ACTRW = Get-StringWithProperty -SearchString $Add -Like 'add rewrite action *';

if ($Null -eq $ACTRW -or $ACTRW.Length -le 0) {
        WriteWordLine 0 0 "No Custom Rewrite Policy has been configured"
        } else {
            ## IB - Use an array of hashtable to store the rows
            [System.Collections.Hashtable[]] $ACTRWH = @();

            ## IB - Iterate over all load balancer bindings (uses new function)
            foreach ($POL in $ACTRW) {
                $Y = Get-StringPropertySplit $POL �RemoveQuotes
                $ACTRWH += @{ 
                    Rewrite = $Y[3]; 
                    Rule = $Y[4];
                    Undefined = $Y[5];
                    }
                } # end foreach

            ## IB - Create the parameters to pass to the AddWordTable function
            $Params = $null
            $Params = @{
                Hashtable = $ACTRWH;
                AutoFit = $wdAutoFitContent;
                Format = -235; ## IB - Word constant for Light List Accent 5
                }
            ## IB - Add the table to the document, splatting the parameters
            $Table = AddWordTable @Params -NoGridLines;
            }
        FindWordDocumentEnd;
        WriteWordLine 0 0 " "

$selection.InsertNewPage()

#endregion NetScaler Actions

#region NetScaler Profiles
$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler Profiles"

WriteWordLine 1 0 "NetScaler Profiles"

WriteWordLine 2 0 "NetScaler Custom TCP Profiles"

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler TCP Profiles Table"

$TCPPROFILES = Get-StringWithProperty -SearchString $Set -Like 'set ns tcpProfile*';

if($TCPPROFILES.Length -le 0) { WriteWordLine 0 0 "No Custom TCP Profiles has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $TCPPROFILESH = @();

    foreach ($TCPPROFILE in $TCPPROFILES) {

        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $TCPPROFILEPropertyArray = Get-StringPropertySplit -SearchString ($TCPPROFILE -Replace 'set ns tcpProfile' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $TCPPROFILESH += @{
                TCP = $TCPPROFILEPropertyArray[0];
                WS = Get-StringProperty $TCPPROFILE "-WS" "NA";
                SACK = Get-StringProperty $TCPPROFILE "-SACK" "NA";
                NAGLE = Get-StringProperty $TCPPROFILE "-NAGLE" "NA";
                MSS = Get-StringProperty $TCPPROFILE "-MSS" "NA";
            }
        }

        if ($TCPPROFILESH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $TCPPROFILESH;
                Columns = "TCP","WS","SACK","NAGLE","MSS";
                Headers = "TCP","WS","SACK","NAGLE","MSS";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }
    
WriteWordLine 2 0 "NetScaler Custom HTTP Profiles"

Write-Verbose "$(Get-Date): `t`tTable: Write NetScaler HTTP Profiles Table"

$HTTPPROFILES = Get-StringWithProperty -SearchString $Add -Like 'add ns httpProfile*';

if($HTTPPROFILES.Length -le 0) { WriteWordLine 0 0 "No Custom HTTP Profiles has been configured"} else {
    
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $HTTPPROFILESH = @();

    foreach ($HTTPPROFILE in $HTTPPROFILES) {
        ## IB - We can now utilise the Get-StringPropertySplit function to return all escaped properties, including quoted expressions
        $HTTPPROFILEPropertyArray = Get-StringPropertySplit -SearchString ($HTTPPROFILE -Replace 'add ns httpProfile' ,'') -RemoveQuotes;
        ## IB - Create parameters for the hashtable so that we can splat them otherwise the
        ## IB - command will be able 400 characters wide!

        $HTTPPROFILESH += @{
                HTTP = $HTTPPROFILEPropertyArray[0];
                Drop = Get-StringProperty $HTTPPROFILE "-dropInvalReqs" "Disabled";
                SPDY = Get-StringProperty $HTTPPROFILE "-spdy" "Disabled";
            }
        }
        if ($HTTPPROFILESH.Length -gt 0) {
            $Params = $null
            $Params = @{
                Hashtable = $HTTPPROFILESH;
                Columns = "HTTP","Drop","SPDY";
                Headers = "HTTP Profile","Drop Invalid Connections","SPDY";
                Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
                AutoFit = $wdAutoFitContent;
                }
            $Table = AddWordTable @Params -NoGridLines;
            FindWordDocumentEnd;
            WriteWordLine 0 0 " "
            }
        }
$selection.InsertNewPage()

#endregion NetScaler Profiles

#endregion NetScaler Documentation Script Complete

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

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
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date -Format G): Trying to send email using current user's credentials without SSL"
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
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

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
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region script template 2

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

###Change the two lines below for your script
$AbstractTitle = "NetScaler Documentation Report"
$SubjectTitle = "NetScaler Documentation Report"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

If($MSWORD -or $PDF)
{
    SaveandCloseDocumentandShutdownWord
}

Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

$GotFile = $False

If($PDF)
{
	If(Test-Path "$($Script:FileName2)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
		$GotFile = $True
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
		Write-Error "Unable to save the output file, $($Script:FileName2)"
	}
}
Else
{
	If(Test-Path "$($Script:FileName1)")
	{
		Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
		$GotFile = $True
	}
	Else
	{
		Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
		Write-Error "Unable to save the output file, $($Script:FileName1)"
	}
}

#email output file if requested
If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
{
	If($PDF)
	{
		$emailAttachment = $Script:FileName2
	}
	Else
	{
		$emailAttachment = $Script:FileName1
	}
	SendEmail $emailAttachment
}

Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null

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
	$SIFile = "$($pwd.Path)\XAXDV2InventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	Out-File -FilePath $SIFile -InputObject "" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $($AddDateTime)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $($Script:CoName)" 4>$Null		
	Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $($CompanyAddress)" 4>$Null		
	Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $($CompanyEmail)" 4>$Null		
	Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $($CompanyFax)" 4>$Null		
	Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $($CompanyPhone)" 4>$Null		
	Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $($CoverPage)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Dev                : $($Dev)" 4>$Null
	If($Dev)
	{
		Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $($Script:DevErrorFile)" 4>$Null
	}
	Out-File -FilePath $SIFile -Append -InputObject "Filename1          : $($Script:FileName1)" 4>$Null
	If($PDF)
	{
		Out-File -FilePath $SIFile -Append -InputObject "Filename2          : $($Script:FileName2)" 4>$Null
	}
	Out-File -FilePath $SIFile -Append -InputObject "Folder             : $($Folder)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "From               : $($From)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $($PDF)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $($MSWORD)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $($ScriptInfo)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $($SmtpPort)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $($SmtpServer)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Title              : $($Script:Title)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "To                 : $($To)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $($UseSSL)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "User Name          : $($UserName)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $($PSCulture)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Word language      : $($Script:WordLanguageValue)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Word version       : $($Script:WordProduct)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Script start       : $($Script:StartTime)" 4>$Null
	Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $($Str)" 4>$Null
}

#V2.10 added
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
$ErrorActionPreference = $SaveEAPreference
#endregion script template 2
