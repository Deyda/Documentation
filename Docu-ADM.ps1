#Requires -Version 4.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
    Creates a complete inventory of a Citrix NetScaler MAS configuration using Microsoft 
	Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix NetScaler MAS configuration using Microsoft 
	Word and PowerShell.
	Creates a Word document named "NetScaler MAS Documentation".
	Document includes a Cover Page, Table of Contents and Footer.
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

	Script requires at least PowerShell version 4 but runs best in version 5.

.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be ReportName_2020-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report 
	will not contain a Company Name on the cover page.
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
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER Credential
	NetScaler MAS username
	
	Specifies a user name for the NetScaler MAS credential, such as "User01". 
	
	Note: for MAS you do not provide a "domain" prefix such as domain\user as the 
	prefix is reserved for tenant identification

	You are prompted for a password.

	If you omit this parameter, you are prompted for a user name and a password.	
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER NSIP
    NetScaler MAS IP address, could be either node in an HA deployment
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER UseNSSSL
	EXPERIMENTAL: Requires SSL/TLS, e.g. https://. 
	This requires the client to trust the NetScaler MAS certificate.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Andrew 
	McCullough" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Andrew 
	McCullough"
	$env:username = Administrator

	Andrew McCullough for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Andrew 
	McCullough" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Andrew 
	McCullough"
	$env:username = Administrator

	Andrew McCullough for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\Docu-ADM.ps1 -CompanyName "Andrew McCullough 
	Consulting" -CoverPage "Mod" -UserName "Andrew McCullough"

	Will use:
		Andrew McCullough Consulting for the Company Name.
		Mod for the Cover Page format.
		Andrew McCullough for the User Name.
.EXAMPLE
	PS C:\PSScript .\Docu-ADM.ps1 -CN "Andrew McCullough 
	Consulting" -CP "Mod" -UN "Andrew McCullough"

	Will use:
		Andrew McCullough Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Andrew McCullough for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1 -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Andrew 
	McCullough" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Andrew 
	McCullough"
	$env:username = Administrator

	Andrew McCullough for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be Script_Template_2020-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Andrew 
	McCullough" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Andrew 
	McCullough"
	$env:username = Administrator

	Andrew McCullough for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be Script_Template_2020-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Andrew 
	McCullough" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Andrew 
	McCullough"
	$env:username = Administrator

	Andrew McCullough for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Docu-ADM.ps1 
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
	PS C:\PSScript > .\Docu-ADM.ps1 
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
	PS C:\PSScript > .\Docu-ADM.ps1 
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
	PS C:\PSScript > .\Docu-ADM.ps1 
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
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word or PDF document.
.NOTES
	NAME: Docu-ADM.ps1
	VERSION: 1.11
	AUTHOR: Andy McCullough, Barry Schiffer, Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters
	LASTEDIT: May 11, 2020
#>

#region changelog
<#
.COMMENT
    If you find issues with saving the final document or table layout is messed up please use the X86 version of Powershell!
.NetScaler MAS Documentation Script
    NAME: Docu-ADM.ps1
	VERSION NetScaler MAS Script: 1.0
	AUTHOR NetScaler MAS script: Andy McCullough
    AUTHOR NetScaler MAS script functions: Andy McCullough, Iain Brighton
    AUTHOR Original NetScaler Documentation Script: Barry Schiffer
    AUTHOR Script template: Carl Webster, Michael B. Smith, Iain Brighton, Jeff Wouters

#Version 1.11 11-May-2020 [all updates made by Webster]
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Add -Log parameter
#	Change color variables $wdColorGray15, $wdColorGray05, and $wdColorRed from [long] to [int]
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Remove the SMTP parameterset and manually verify the parameters
#	Reorder parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
#	Update Function SetWordCellFormat to change parameter $BackgroundColor to [int]
#	Update Help Text
#

#Version 1.1 14-Sep-2018
# Fixed issue with Stylebook Content being output in Base64
# Fixed issue with Stylebook description containing newline characters breaking tables

#Version 1.01 13-Feb-2017
#	Fixed French wording for Table of Contents 2 (Thaqnks to David Rouquier)

#Version 1.0 1-Dec-2016
#	This version includes:
#
#	MAS System Configuration
#	Basic Configuration
#	System Administration Settings
#	Licensing
#	Notification Settings
#	SNMP Configuration
#	Authentication Settings
#	Device Profiles
#	Managed Instances
#	Instance Groups
#	Event Management
#	Configuration Templates
#	DataCenters/IP Blocks
#	Stylebooks
#	Analytics Settings
#
#>
#endregion changelog

#region script template
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False )] 
	[Switch]$AddDateTime=$False,
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

    [parameter(Mandatory=$false ) ]
    [PSCredential] $Credential = (Get-Credential -Message 'Enter NetScaler MAS credentials'),
	
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
    [parameter(Mandatory=$true )]
    [Alias("MASIP")]
    [string] $NSIP,
    
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	## EXPERIMENTAL: Requires SSL/TLS, e.g. https://. This requires the client to trust the NetScaler MAS certificate.
    [parameter(Mandatory=$false )]
	[System.Management.Automation.SwitchParameter] $UseNSSSL,
    
	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To=""

	)

Set-StrictMode -Version Latest

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

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
	Write-Error "
	`n`n
	`t`t
	Unable to determine output parameter.
	`n`n
	`t`t
	Script cannot continue.
	`n`n
	"
	Exit
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
			Exit
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
		Exit
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
	$Script:LogPath = "$Script:pwdpath\NSMASInventoryScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
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
	$Script:DevErrorFile = "$Script:pwdpath\NSMASInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
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

#region initialize variables for word
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

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
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

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
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
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
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
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
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
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
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
		Write-Error "
		`n`n
		`t`t
		The Word object could not be created.  You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
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
		Write-Error "
		`n`n
		`t`t
		Unable to determine the Word language value.
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
		$ErrorActionPreference = $SaveEAPreference
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
		$ErrorActionPreference = $SaveEAPreference
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

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
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
		Write-Error "
		`n`n
		`t`t
		An empty Word document could not be created.
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
		$ErrorActionPreference = $SaveEAPreference
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
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

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

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
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

#region word output functions
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
#endregion

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
	is returned).
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
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
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
		} ## end ElseIf
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
					## Build the available columns from all available PSCustomObject note properties
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
				} ## end foreach
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
				} ## end foreach

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
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
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
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
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
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
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
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
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
Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((gm -Name $topLevel -InputObject $object))
		{
			If((gm -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): AddDateTime     : $($AddDateTime)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
	}
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): Dev             : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Filename1       : $($Script:filename1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $($Script:filename2)"
	}
	Write-Verbose "$(Get-Date): Folder          : $($Folder)"
	Write-Verbose "$(Get-Date): From            : $($From)"
	Write-Verbose "$(Get-Date): NSIP            : $($NSIP)"
	Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
	Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
	Write-Verbose "$(Get-Date): To              : $($To)"
	Write-Verbose "$(Get-Date): Use NS SSL      : $($UseNSSSL)"
	Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name       : $($UserName)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language   : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version    : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start    : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
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

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($Script:pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($Script:pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
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
		$SIFile = "$Script:pwdpath\NSMASInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime   : $($AddDateTime)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name   : $($Script:CoName)" 4>$Null		
		}
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page     : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev            : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile   : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1      : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2      : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder         : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From           : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "NSIP           : $($NSIP)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF    : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD   : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info    : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port      : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server    : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title          : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To             : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use NS SSL     : $($UseNSSSL)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL        : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name      : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected    : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version   : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture      : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture    : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language  : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version   : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start   : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time   : $($Str)" 4>$Null
	}

	$ErrorActionPreference = $SaveEAPreference
	[gc]::collect()
}
#endregion

#region general script functions
Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}

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
	[gc]::collect()
}
#endregion


#Script begins

$script:startTime = Get-Date
#endregion script template

#region file name and title name
#The function SetFileName1andFileName2 needs your script output filename
#change title for your report
[string]$Script:Title = "NetScaler MAS Documentation V1.0 $($Script:CoName)"
SetFileName1andFileName2 "NetScaler MAS Documentation V1.0"

#endregion file name and title name

#region NetScaler Documentation Script Complete

## Barry Schiffer Use Stopwatch class to time script execution
$sw = [Diagnostics.Stopwatch]::StartNew()

$selection.InsertNewPage()

#region Nitro Functions

function Get-vNetScalerObjectList {
<#
    .SYNOPSIS
        Returns a list of objects available in a NetScaler Nitro API container.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Container, i.e. nitro/v2/stat/ or nitro/v2/config/
        [Parameter(Mandatory)] [ValidateSet('Stat','Config')] [System.String] $Container
    )
    begin {
        $Container = $Container.ToLower();
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/nitro/v2/{2}/' -f $protocol, $script:nsSession.Address, $Container;
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        $methodResponse = '{0}objects' -f $Container.ToLower();
        Write-Output $restResponse.($methodResponse).objects;
    }
} #end function Get-vNetScalerObjectList

function Get-vNetScalerObject {
<#
    .SYNOPSIS
        Returns a NetScaler Nitro API object(s) via its REST API.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API resource type, e.g. /nitro/v2/config/LBVSERVER
        [Parameter(Mandatory)] [Alias('Object','Type')] [System.String] $ResourceType,
        # NetScaler Nitro API resource name, e.g. /nitro/v2/config/lbvserver/MYLBVSERVER
        [Parameter()] [Alias('Name')] [System.String] $ResourceName,
        # NetScaler Nitro API optional attributes, e.g. /nitro/v2/config/lbvserver/mylbvserver?ATTRS=<attrib1>,<attrib2>
        [Parameter()] [System.String[]] $Attribute,
        # NetScaler Nitro API Container, i.e. nitro/v2/stat/ or nitro/v2/config/
        [Parameter()] [ValidateSet('Stat','Config')] [System.String] $Container = 'Config'
    )
    begin {
        $Container = $Container.ToLower();
        $ResourceType = $ResourceType.ToLower();
        $ResourceName = $ResourceName.ToLower();
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/nitro/v1/{2}/{3}' -f $protocol, $script:nsSession.Address, $Container, $ResourceType;
        if ($ResourceName) { $uri = '{0}/{1}' -f $uri, $ResourceName; }
        if ($Attribute) {
            $attrs = [System.String]::Join(',', $Attribute);
            $uri = '{0}?attrs={1}' -f $uri, $attrs;
        }
        $uri = [System.Uri]::EscapeUriString($uri.ToLower());
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        if ($null -ne $restResponse.($ResourceType)) { Write-Output $restResponse.($ResourceType); }
        else { Write-Output $restResponse }
    }
}
 #end function Get-vNetScalerObject

function Get-vNetScalerStylebookObject {
<#
    .SYNOPSIS
        Returns a NetScaler Nitro API object(s) via its REST API.
#>
    [CmdletBinding()]
    param (
        # Stylebook Namespace Value
        [Parameter(Mandatory)] [System.String] $NameSpace,
        # Stylebook Version Value
        [Parameter(Mandatory)] [System.String] $Version,
        # Stylebook Name Value
        [Parameter(Mandatory)] [System.String] $Name,
        # NetScaler Nitro API Container, i.e. nitro/v2/stat/ or nitro/v2/config/
        [Parameter()] [ValidateSet('Stat','Config')] [System.String] $Container = 'Config'
    )
    begin {
        $Container = $Container.ToLower();
        $NameSpace = $NameSpace.ToLower();
   
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/stylebook/nitro/v1/{2}/stylebooks/{3}/{4}/{5}' -f $protocol, $script:nsSession.Address, $Container, $NameSpace, $Version, $Name;
        
        $uri = [System.Uri]::EscapeUriString($uri);
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;

        if ($null -ne $restResponse.stylebook) { 
            Write-Output $restResponse.stylebook; } else { 
            Write-Output $restResponse 
        }
}
}

     #end function Get-vNetScalerStylebookObject

     function Get-vNetScalerStylebooks {
<#
    .SYNOPSIS
        Returns a NetScaler Nitro API object(s) via its REST API.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Container, i.e. nitro/v2/stat/ or nitro/v2/config/
        [Parameter()] [ValidateSet('Stat','Config')] [System.String] $Container = 'Config'
    )
    begin {
        $Container = $Container.ToLower();
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/stylebook/nitro/v1/{2}/stylebooks' -f $protocol, $script:nsSession.Address, $Container;
        
        $uri = [System.Uri]::EscapeUriString($uri.ToLower());
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;

        if ($null -ne $restResponse.stylebooks) { 
            Write-Output $restResponse.stylebooks; } else { 
            Write-Output $restResponse 
        }
}
}

     #end function Get-vNetScalerStylebooks

function Get-vNetScalerFile {

<#
    .SYNOPSIS
        Returns a NetScaler Nitro API SystemFile object(s) via its REST API.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API resource name, e.g. /nitro/v2/config/SystemFile?args=filename:Filename,filelocation:FileLocation
        [Parameter()] [Alias('Name')][System.String] $FileName,
        # NetScaler Nitro API optional attributes, e.g. /nitro/v2/config/lbvserver/mylbvserver?ATTRS=<attrib1>,<attrib2>
        [Parameter()] [Alias('Location')] [System.String] $FileLocation
        
    )
    begin {
        #Don't lower case these as they are case sensitive
        #$FileName = $FileName.ToLower();
        $FileLocation = $FileLocation.Replace("/","%2F");
        $Container = "config"
        
    }
    process {
        if ($script:nsSession.UseSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $uri = '{0}://{1}/nitro/v2/config/systemfile/{2}?args=filelocation:{3}' -f $protocol, $script:nsSession.Address, $FileName, $FileLocation;
        
        #Don't URI encode as we've already replaced / with %2F as required - URL encoding after this, encodes the % which breaks the request
        #$uri = [System.Uri]::EscapeUriString($uri);
        #Write-Output $uri;
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        $test = 1
        if ($null -ne $restResponse.systemfile) { Write-Output $restResponse.systemfile; }
        else { Write-Output $restResponse }
    }
} #end function Get-vNetScalerFile

function InvokevNetScalerNitroMethod {
<#
    .SYNOPSIS
        Calls a fully qualified NetScaler Nitro API
    .NOTES
        This is an internal function and shouldn't be called directly
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API uniform resource identifier
        [Parameter(Mandatory)] [string] $Uri,
        # NetScaler Nitro API Container, i.e. nitro/v2/stat/ or nitro/v2/config/
        [Parameter(Mandatory)] [ValidateSet('Stat','Config')] [string] $Container
    )
    begin {
        if ($script:nsSession -eq $null) { throw 'No valid NetScaler session configuration.'; }
        if ($script:nsSession.Session -eq $null -or $script:nsSession.Expiry -eq $null) { throw 'Invalid NetScaler session cookie.'; }
        if ($script:nsSession.Expiry -lt (Get-Date)) { throw 'NetScaler session has expired.'; }
    }
    process {
        $irmParameters = @{
            Uri = $Uri;
            Method = 'Get';
            WebSession = $script:nsSession.Session;
            ErrorAction = 'Stop';
            Verbose = ($PSBoundParameters['Debug'] -eq $true);
        }
        Write-Output (Invoke-RestMethod @irmParameters);
    }
} #end function InvokevNetScalerNitroMethod

function Connect-vNetScalerSession {
<#
    .SYNOPSIS
        Authenticates to the NetScaler and stores a session cookie.
#>
    [CmdletBinding(DefaultParameterSetName='HTTP')]
    [OutputType([Microsoft.PowerShell.Commands.WebRequestSession])]
    param (
        # NetScaler uniform resource identifier
        [Parameter(Mandatory, ParameterSetName='HTTP')]
        [Parameter(Mandatory, ParameterSetName='HTTPS')]
        [System.String] $ComputerName,
        # NetScaler session timeout (seconds)
        [Parameter(ParameterSetName='HTTP')]
        [Parameter(ParameterSetName='HTTPS')]
        [ValidateNotNull()]
        [System.Int32] $Timeout = 3600,
        # NetScaler authentication credentials
        [Parameter(ParameterSetName='HTTP')]
        [Parameter(ParameterSetName='HTTPS')]
        [System.Management.Automation.PSCredential] $Credential = $(Get-Credential -Message "Provide NetScaler credentials for '$ComputerName'";),
        ## EXPERIMENTAL: Require SSL/TLS, e.g. https://. This requires the client to trust to the NetScaler's certificate.
        [Parameter(ParameterSetName='HTTPS')] [System.Management.Automation.SwitchParameter] $UseNSSSL
    )
    process {
        if ($UseNSSSL) { $protocol = 'https'; }
        else { $protocol = 'http'; }
        $script:nsSession = @{ Address = $ComputerName; UseSSL = $UseNSSSL }
        $json = '{{"login":{{"username":"{0}","password":"{1}","session_timeout":{2}}}}}' -f $Credential.UserName, $Credential.GetNetworkCredential().Password, $Timeout;
        $jsonuri = [System.Net.WebUtility]::UrlEncode($json);
        $objjson = "object=$jsonuri";
        $invokeRestMethodParams = @{
            Uri = ('{0}://{1}/nitro/v1/config/login' -f $protocol, $ComputerName);
            Method = 'Post';
            Body = ($objjson);
            ContentType = 'application/x-www-form-urlencoded';
            SessionVariable = 'nsSessionCookie';
            ErrorAction = 'Stop';
        }
        $restResponse = Invoke-RestMethod @invokeRestMethodParams;
        ## Store the session cookie at the script scope
        $script:nsSession.Session = $nsSessionCookie;
        ## Store the session expiry
        $script:nsSession.Expiry = (Get-Date).AddSeconds($Timeout);
        ## Return the Rest Method response
        Write-Output $restResponse;
    }
} #end function Connect-vNetScalerSession

function Get-vNetScalerObjectCount {
<#
.Synopsis
    Returns an individual NetScaler Nitro API object.
#>
    [CmdletBinding()]
    param (
        # NetScaler Nitro API Object, e.g. /nitro/v2/config/NSVERSION
        [Parameter(Mandatory)] [string] $Object,
        # NetScaler Nitro API Container, i.e. nitro/v2/stat/ or nitro/v2/config/
        [Parameter(Mandatory)] [ValidateSet('Stat','Config')] [string] $Container
    )

    begin {
        ## Check session cookie
        if ($script:nsSession.Session -eq $null) { throw 'Invalid NetScaler session cookie.'; }
    }

    process {
        $uri = 'http://{0}/nitro/v1/{1}/{2}?count=yes' -f $script:nsSession.Address, $Container.ToLower(), $Object.ToLower();
        $restResponse = InvokevNetScalerNitroMethod -Uri $Uri -Container $Container;
        # $objectResponse = '{0}objects' -f $Container.ToLower();
        Write-Output $restResponse.($Object.ToLower());
    }
}

#endregion Nitro Functions


#region generic functions

function IsNull($objectToCheck) {
    if ($objectToCheck -eq $null) {
        return $true
    }

    if ($objectToCheck -is [String] -and $objectToCheck -eq [String]::Empty) {
        return $true
    }

    if ($objectToCheck -is [DBNull] -or $objectToCheck -is [System.Management.Automation.Language.NullString]) {
        return $true
    }

    return $false
}

function GetMASLicense($objectToCheck) {
    if ($objectToCheck -eq 0) {
        return "Unlicensed/Disabled"
    }

    if ($objectToCheck -eq 1) {
        return "Unlicensed/Enabled"
    }

    if ($objectToCheck -eq 2) {
        return "Licensed/Disabled"
    }

    if ($objectToCheck -eq 3) {
        return "Licensed/Enabled"
    }


    return "Unknown License State"
}

function GetDeviceFamily($objFamily){

 switch($objFamily.ToLower()) {
    "cbwanopt" {return "CloudBridge WANOp"}
    "nssdx" {return "NetScaler SDX"}
    "cb" {return "CloudBridge"}
    "ns" {return "NetScaler"}
    "nsvpx" {return "NetScaler VPX"}
    default {return $objFamily}
 }
}

 function GetDuration($objDuration){

 switch($objDuration.ToLower()) {
    "l3" {return "Hour"}
    "l4" {return "Day"}
    "l5" {return "Weekly"}
    default {return $objDuration}
 }
 }

function GetPlatform($objPlatform) {

switch($objPlatform.ToLower()) {
    "450000" {return "VPX on XenServer"}
    "450010" {return "VPX on VMware ESX"}
    "450020" {return "VPX on Microsoft Hyper-V"}
    "450070" {return "VPX on generic KVM"}
    "450040" {return "VPX on Amazon Web Services"}
    default {return $objPlatform}
 }

}


#endregion generic functions

#region NetScaler MAS Connect

## Ensure we can connect to the NetScaler appliance before we spin up Word!
## Connect to the API if there is no session cookie
## Note: repeated logons will result in 'Connection limit to cfe exceeded' errors.
if (-not (Get-Variable -Name nsSession -Scope Script -ErrorAction SilentlyContinue)) { 
    #[ref] $null = Connect-vNetScalerSession -ComputerName $nsip -Credential $Credential -UseSSL:$UseNSSSL -ErrorAction Stop;
    [ref] $null = Connect-vNetScalerSession -ComputerName $nsip -Credential $Credential -ErrorAction Stop;
}
#endregion NetScaler Connect

#region NetScaler chaptercounters
$Chapters = 38
$Chapter = 0
#endregion NetScaler chaptercounters


#region NetScaler MAS System Information
WriteWordLine 1 0 "NetScaler MAS System Configuration"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Configuration"

#region MAS Basic Configuration
WriteWordLine 2 0 "NetScaler MAS Basic Configuration"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Basic Configuration"

$nmasconfig = Get-vNetScalerObject -Container Config -Object mps

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasconfigh = @(
    @{ Description = "Hostname"; Value = $nmasconfig.hostname ;}
	@{ Description = "Product Build Number"; Value = $nmasconfig.product_build_number ;}
	@{ Description = "Is a member of HA Group"; Value = $nmasconfig.is_member_of_default_group ;}
    @{ Description = "Cloud Deployment"; Value = $nmasconfig.is_cloud ;}
    @{ Description = "Platform"; Value = GetPlatform($nmasconfig.platform) ;}
    @{ Description = "Time Zone"; Value = $nmasconfig.time_zone ;}
    @{ Description = "Product Name"; Value = $nmasconfig.product ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasconfigh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "

    
#endregion MAS Basic Configuration

#region MAS System Administration

WriteWordLine 2 0 "NetScaler MAS System Administration"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Administration"

#region MAS System Settings
WriteWordLine 3 0 "NetScaler MAS System Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Settings"

$nmassyssettings = Get-vNetScalerObject -Container Config -Object system_settings;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasconfigh = @(
    @{ Description = "Communications with instances"; Value = $nmassyssettings.svm_ns_comm ;}
	@{ Description = "Secure Access Only"; Value = $nmassyssettings.secure_access_only ;}
	@{ Description = "Enable Session Timeout"; Value = $nmassyssettings.enable_session_timeout ;}
    @{ Description = "Allow basic authentication"; Value = $nmassyssettings.basicauth ;}
    @{ Description = "Enable nsrecover login"; Value = $nmassyssettings.enable_nsrecover_login ;}
    @{ Description = "Enable Certificate Download"; Value = $nmassyssettings.enable_certificate_download ;}
    @{ Description = "Enable Shell access for non-nsroot users"; Value = $nmassyssettings.enable_shell_access ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasconfigh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "
#endregion MAS System Settings

#region MAS SSL Settings

WriteWordLine 3 0 "NetScaler MAS SSL Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SSL Settings"

$nmasssslsettings = Get-vNetScalerObject -Container Config -Object ssl_settings;
$nmassciphersettings = Get-vNetScalerObject -Container Config -Object cipher_config;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasconfigh = @(
    @{ Description = "Enable SSL Renegotiation"; Value = $nmasssslsettings.sslreneg ;}
    @{ Description = "SSLv3 Enabled"; Value = $nmasssslsettings.sslv3 ;}
    @{ Description = "TLSv1 Enabled"; Value = $nmasssslsettings.tlsv1 ;}
    @{ Description = "TLSv1.1 Enabled"; Value = $nmasssslsettings.tlsv1_1 ;}
    @{ Description = "TLSv1.2 Enabled"; Value = $nmasssslsettings.tlsv1_2 ;}
    @{ Description = "Applied Cipher Suites"; Value = $nmassciphersettings.cipher_name_list_array -Join ", " ;}
	
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasconfigh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "

#endregion MAS SSL Settings

#region MAS System Prune Settings

WriteWordLine 3 0 "NetScaler MAS System Prune Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Prune Settings"

$nmasprune = Get-vNetScalerObject -Container Config -Object prune_policy;


## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasbach = @(
    @{ Description = "Data to keep (Days)"; Value = $nmasprune.data_to_keep_in_days ;}
	
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmaspruneh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "

#endregion MAS System Prune Settings

#region MAS System Backup Settings

WriteWordLine 3 0 "NetScaler MAS System Backup Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Backup Settings"

$nmasbackup = Get-vNetScalerObject -Container Config -Object backup_policy;


## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasbackuph = @(
    @{ Description = "Previous backups to retain"; Value = $nmasbackup.backup_to_retain ;}
    @{ Description = "Encrypt backup file"; Value = $nmasbackup.encrypt_backup_file ;}
    @{ Description = "Enable External Transfer"; Value = $nmasbackup.enable_external_transfer ;}
	
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasbackuph;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "

#endregion MAS System Backup Settings

#region MAS Device Backup Settings

WriteWordLine 3 0 "NetScaler MAS Instance Backup Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Instance Backup Settings"

$nmasinstbackup = Get-vNetScalerObject -Container Config -Object device_backup_policy;


## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasinstbackuph = @(
    @{ Description = "Frequency of Device Backup (Hours)"; Value = $nmasinstbackup.polling_interval ;}
    @{ Description = "Previous backups to retain"; Value = $nmasinstbackup.number_of_backups ;}
    @{ Description = "Encrypt backup file"; Value = $nmasinstbackup.encrypt_backup_file ;}
    
	
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasinstbackuph;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "

#endregion MAS Device Backup Settings

#endregion MAS System Administration


#region MAS Licensing
WriteWordLine 2 0 "NetScaler MAS Licensing"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Licensing"

$nmaslic = Get-vNetScalerObject -Container Config -Object mas_license

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmaslich = @(
    @{ Description = "CPX Licensing"; Value = GetMASLicense($nmaslic.cpx_lic) ;}
	@{ Description = "Analytics"; Value = GetMASLicense($nmaslic.analytics) ;}
    @{ Description = "Advanced Analytics"; Value = GetMASLicense($nmaslic.adv_analytics) ;}
    @{ Description = "Performance"; Value = GetMASLicense($nmaslic.perf) ;}
    @{ Description = "Syslog"; Value = GetMASLicense($nmaslic.syslog) ;}
    @{ Description = "Pooled Licensing"; Value = GetMASLicense($nmaslic.pooled_lic) ;}
    @{ Description = "SNMP Traps"; Value = GetMASLicense($nmaslic.snmp_traps) ;}
	@{ Description = "Maximum VIPs Licensed"; Value = $nmaslic.max_vips ;}
    @{ Description = "Total managed VIPs"; Value = $nmaslic.total_managed_vips ;}
    @{ Description = "Total Discovered VIPs"; Value = $nmaslic.total_discovered_vips ;}
    @{ Description = "Total Allowed VIPs"; Value = $nmaslic.total_allowed_vips ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmaslich;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null 
WriteWordLine 0 0 " "     

#endregion MAS Licensing


#region MAS Notifications
WriteWordLine 2 0 "NetScaler MAS Notifications"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Notifications"

#region SMTP Servers
WriteWordLine 3 0 "SMTP Servers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SMTP Servers"

$nmassmtp = Get-vNetScalerObject -Container Config -Object smtp_server

If (!$nmassmtp) { WriteWordLine "No SMTP Servers have been configured" } Else { 
Foreach ($nmassmtpserver in $nmassmtp) {
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmassmtph = @(
    @{ Description = "Server Name/IP"; Value = $nmassmtpserver.server_name ;}
    @{ Description = "Port"; Value = $nmassmtpserver.port ;}
    @{ Description = "Use Authentication"; Value = $nmassmtpserver.is_auth ;}
    @{ Description = "Username"; Value = $nmassmtpserver.username ;}
    @{ Description = "Password"; Value = $nmassmtpserver.password ;}
    @{ Description = "Secure"; Value = $nmassmtpserver.is_ssl ;}
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmassmtph;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;
WriteWordLine 0 0 " "
FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null 
} #end foreach smtp server
} #end if
WriteWordLine 0 0 " "     

#endregion SMTP Servers

#region SMTP Distribution Groups 

WriteWordLine 3 0 "SMTP Distribution Groups"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SMTP Distribution Groups"

$nmassmtpgrps = Get-vNetScalerObject -Container Config -Object mail_profile

If (!$nmassmtpgrps) { WriteWordLine "No SMTP Distribution Groups have been configured" } Else { 
Foreach ($nmassmtpgrp in $nmassmtpgrps) {
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmassmtpgrph = @(
    @{ Description = "Group Name"; Value = $nmassmtpgrp.profile_name ;}
    @{ Description = "Server Name/IP"; Value = $nmassmtpgrp.server_name ;}
    @{ Description = "Sender Address"; Value = $nmassmtpgrp.sender_mail_address ;}
    @{ Description = "To List"; Value = $nmassmtpgrp.to_list ;}
    @{ Description = "CC List"; Value = $nmassmtpgrp.cc_list ;}
    @{ Description = "BCC List"; Value = $nmassmtpgrp.bcc_list ;}
    
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmassmtpgrph;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null 
$nmassmtpgrph = $Null
WriteWordLine 0 0 " "
} #end foreach smtp server
} #end if
WriteWordLine 0 0 " "  


#endregion SMTP Distribution Groups 

#region SMS Servers
WriteWordLine 3 0 "SMS Servers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SMS Servers"

$nmassmssrvs = Get-vNetScalerObject -Container Config -Object sms_server

Foreach ($nmassmssrv in $nmassmssrvs) {
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmassmssrvh = @(
    @{ Description = "Server Name/IP"; Value = $nmassmssrv.server_name ;}
    @{ Description = "Username Key"; Value = $nmassmssrv.username_key ;}
    @{ Description = "Username Value"; Value = $nmassmssrv.username_val ;}
    @{ Description = "Password Key"; Value = $nmassmssrv.password_key ;}
    @{ Description = "Password Value"; Value = $nmassmssrv.password_val ;}
    @{ Description = "Optional Key"; Value = $nmassmssrv.optional1_key ;}
    @{ Description = "Optional Value"; Value = $nmassmssrv.optional1_val ;}
    @{ Description = "Base URL"; Value = $nmassmssrv.base_url ;}
    @{ Description = "To Key"; Value = $nmassmssrv.to_key ;}
    @{ Description = "To Separator"; Value = $nmassmssrv.to_seperater ;}
    @{ Description = "Message Key"; Value = $nmassmssrv.message_key ;}
    @{ Description = "Message Word Separator"; Value = $nmassmssrv.message_word_sperater ;}
    @{ Description = "Request Type"; Value = $nmassmssrv.type ;}
    @{ Description = "Secure"; Value = $nmassmssrv.is_ssl ;}
    
);

## IB - Create the parameters to pass to the AddWordTable function
If ($nmassmssrvh.Length -gt 0) {
$Params = $null
$Params = @{
    Hashtable = $nmassmssrvh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;
} Else {
      WriteWordLine 0 0 " "
      WriteWordLine 0 0 "No SMS Servers have been configured."
}
WriteWordLine 0 0 " "
FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null 
} #end foreach sms server

WriteWordLine 0 0 " "     

#endregion SMS Servers

#region SMS Servers
WriteWordLine 3 0 "SMS Distribution Lists"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SMS Distribution Lists"

$nmassmsgrps = Get-vNetScalerObject -Container Config -Object sms_server


Foreach ($nmassmsgrp in $nmassmsgrps) {
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmassmsgrph = @(
    @{ Description = "List Name"; Value = $nmassmsgrp.profile_name ;}
    @{ Description = "SMS Server Name"; Value = $nmassmsgrp.server_name ;}
    @{ Description = "To List"; Value = $nmassmsgrp.to_list ;}
    
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmassmssrvh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null 
WriteWordLine 0 0 " "
} #end foreach sms distro group


WriteWordLine 0 0 " "     

#endregion SMS Servers

#endregion MAS Notifications

#region SNMP
WriteWordLine 2 0 "NetScaler MAS SNMP"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SNMP"

#region SNMP Managers
WriteWordLine 3 0 "SNMP Managers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS SNMP Managers"

$nmassnmpmgrs = Get-vNetScalerObject -Container Config -Object snmp_manager

    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasnmpmgrh = @();

foreach ($nmassnmpmgr in $nmassnmpmgrs) {

If (IsNull($nmassnmpmgr.netmask)){$nmassnmpmgrmask = ""} Else { $nmassnmpmgrmask = $nmassnmpmgr.netmask};


    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $nmasnmpmgrh += @{
            ManagerName = $nmassnmpmgr.ip_address;
            NetMask = $nmassnmpmgrmask;
            Community = $nmassnmpmgr.community;
        }
    }

if ($nmasnmpmgrh.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $nmasnmpmgrh;
        Columns = "ManagerName","NetMask","Community";
        Headers = "Manager Name/IP Address", "NetMask", "Community String";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null
    } Else {
      WriteWordLine 0 0 " "
      WriteWordLine 0 0 "No SNMP Managers have been configured."
      WriteWordLine 0 0 " "
    }

#endregion SNMP Managers


#endregion SNMP

#region authentication
WriteWordLine 2 0 "NetScaler MAS Authentication"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Authentication"

#region MAS Users
WriteWordLine 3 0 "NetScaler MAS System Users"
WriteWordLine 0 0 " "


$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Users"

$nmasusers = Get-vNetScalerObject -Container Config -Object mpsuser

  ForEach ($nmasuser in $nmasusers) {
    $nmasusername = $nmasuser.name
    WriteWordLine 4 0 "System User: $nmasusername"
    WriteWordLine 0 0 " "

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $nmasuserh = @(
    @{ Description = "Name"; Value = $nmasuser.name ;}
    @{ Description = "Permissions"; Value = $nmasuser.permission ;}
    @{ Description = "Session Timeout"; Value = $nmasuser.session_timeout ;}
    @{ Description = "Session Timeout Unit"; Value = $nmasuser.session_timeout_unit ;}
    @{ Description = "Enable External Authentication"; Value = $nmasuser.external_authentication ;}
    @{ Description = "Tenant Name"; Value = $nmasuser.tenant_name ;}
    @{ Description = "Groups"; Value = $nmasuser.groups ;}

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasuserh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "


  }




#endregion MAS Users


#region MAS Groups
WriteWordLine 3 0 "NetScaler MAS System Groups"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS System Groups"

$nmasgroups = Get-vNetScalerObject -Container Config -Object mpsgroup

If (!$nmasgroups) {
  WriteWordLine 0 0 "No NMAS System Groups have been configured."
  WriteWordLine 0 0 " "
} Else {

  ForEach ($nmasgroup in $nmasgroups) {
    $nmasgroupname = $nmasgroup.name
    WriteWordLine 4 0 "System Group: $nmasgroupname"
    WriteWordLine 0 0 " "

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $nmasgrouph = @(
    @{ Description = "Name"; Value = $nmasgroup.name ;}
    @{ Description = "Permission"; Value = $nmasgroup.permission ;}
    @{ Description = "Session Timeout"; Value = $nmasgroup.session_timeout ;}
    @{ Description = "Session Timeout Enabled"; Value = $nmasgroup.enable_session_timeout ;}
    @{ Description = "Session Timeout Unit"; Value = $nmasgroup.session_timeout_unit ;}
    @{ Description = "Assign All Devices"; Value = $nmasgroup.assign_all_devices ;}
    @{ Description = "Assign All Applications"; Value = $nmasgroup.assign_all_applications ;}
    @{ Description = "Allow Applications Only"; Value = $nmasgroup.allow_application_only ;}
    @{ Description = "Users"; Value = $nmasgroup.users -Join ", " ;}

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasgrouph;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "


  } #end foreach

} #end if

#endregion MAS Groups

#region Authentication Configuration

WriteWordLine 3 0 "NetScaler MAS Authentication Configuration"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Authentication COnfiguration"

$nmasauthconfig = Get-vNetScalerObject -Container Config -Object aaa_server

If ($nmasauthconfig.primary_server_type -eq "LOCAL") {$nmasauthsource = "LOCAL"} Else {$nmasauthsource = "EXTERNAL"};

    [System.Collections.Hashtable[]] $nmasauthconfigh = @(
    @{ Description = "Authentication Source"; Value = $nmasauthsource ;}
    @{ Description = "Enable Fallback to Local Authentication"; Value = $nmasauthconfig.fallback_local_authentication ;}

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasauthconfigh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "

If ($nmasauthsource -ne "LOCAL") {

#External Auth Servers

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasextauthsrvsh = @();

#Do Primary Ext Server

$nmasextauthsrvsh += @{
            Priority = "1";
            Name = $nmasauthconfig.primary_server_name;
            AuthType = $nmasauthconfig.primary_server_name;
        }

foreach ($nmasextauthsrv in $nmasauthconfig.external_servers) {



    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $nmasextauthsrvsh += @{
            Priority = $nmasextauthsrv.priority;
            Name = $nmasextauthsrv.external_server_name;
            AuthType = $nmasextauthsrv.external_server_type;
            
        }
    }

    $Params = $null
    $Params = @{
        Hashtable = $nmasextauthsrvsh;
        Columns = "Priority","Name","AuthType";
        Headers = "Priority", "Name", "Authentication Type";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null


}


#endregion Authentication Configuration

#region Authentication Servers

WriteWordLine 3 0 "NetScaler MAS Authentication Servers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Authentication Servers"

#region LDAP Servers

WriteWordLine 4 0 "LDAP Authentication Servers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS LDAP Authentication Servers"

$authactsldap = Get-vNetScalerObject -Container config -Object ldap_server;
If (!$authactsldap) {
 WriteWordLine 0 0 "There are no LDAP authentication servers configured. "
}


foreach ($authactldap in $authactsldap) {
    $ACTNAMELDAP = $authactldap.name
    WriteWordLine 5 0 "LDAP Server: $ACTNAMELDAP";
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $LDAPCONFIG = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "LDAP Server IP"; Value = $authactldap.ip_address; }
    @{ Description = "LDAP Server Port"; Value = $authactldap.port; }
    @{ Description = "LDAP Server Time-Out"; Value = $authactldap.auth_timeout; }
    @{ Description = "Validate Certificate"; Value = $authactldap.validate_ldap_server_certs; }
    @{ Description = "LDAP Base OU"; Value = $authactldap.base_dn; }
    @{ Description = "LDAP Bind DN"; Value = $authactldap.bind_dn; }
    @{ Description = "Login Name"; Value = $authactldap.login_name; }
    @{ Description = "Security Type"; Value = $authactldap.sec_type; }   
    @{ Description = "Password Changes"; Value = $authactldap.change_password; }
    @{ Description = "Group attribute name"; Value = $authactldap.group_attr_name; }
    @{ Description = "LDAP Referrals"; Value = $authactldap.max_ldap_referrals; }
    @{ Description = "Nested Group Extraction"; Value = $authactldap.nested_group_extraction; }
    @{ Description = "Maximum Nesting level"; Value = $authactldap.max_nesting_level; }
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
    $Table = AddWordTable @Params -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    WriteWordLine 0 0 " "
}

WriteWordLine 0 0 " "

#endregion LDAP Servers

#region RADIUS Servers

WriteWordLine 4 0 "RADIUS Authentication Servers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS RADIUS Authentication Servers"

$authactsradius = Get-vNetScalerObject -Container config -Object radius_server;
If (!$authactsradius) {
  WriteWordLine 0 0 "There are no RADIUS servers configured."
}
foreach ($authactradius in $authactsradius) {
    $ACTNAMERADIUS = $authactradius.name
    WriteWordLine 5 0 "RADIUS Server: $ACTNAMERADIUS";
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $RADUIUSCONFIG = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "RADIUS Server IP"; Value = $authactradius.ip_address; }
    @{ Description = "RADIUS Server Port"; Value = $authactradius.port; }
    @{ Description = "RADIUS Server Time-Out"; Value = $authactradius.auth_timeout; }
    @{ Description = "Radius NAS IP"; Value = $authactradius.nas_id; }
    @{ Description = "IP Vendor ID"; Value = $authactradius.ip_vendor_id; }
    @{ Description = "Accounting"; Value = $authactradius.accounting; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $RADUIUSCONFIG;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    WriteWordLine 0 0 " "
}

WriteWordLine 0 0 " "

#endregion RADIUS Servers

#region TACACS Servers
WriteWordLine 4 0 "TACACS Authentication Servers"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS TACACS Authentication Servers"

$authactstacacs = Get-vNetScalerObject -Container config -Object tacacs_server;
If (!$authactstacacs) {
  WriteWordLine 0 0 "There are no TACACS servers configured."
}
foreach ($authacttacacs in $authactstacacs) {
    $ACTNAMETACACS = $authacttacacs.name
    WriteWordLine 5 0 "TACACS Server: $ACTNAMETACACS";
    WriteWordLine 0 0 " "
    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $TACACSCONFIG = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "TACACS Server IP"; Value = $authacttacacs.ip_address; }
    @{ Description = "TACACS Server Port"; Value = $authacttacacs.port; }
    @{ Description = "TACACS Server Time-Out"; Value = $authacttacacs.auth_timeout; }
    @{ Description = "Accounting"; Value = $authacttacacs.accounting; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $TACACSCONFIG;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    WriteWordLine 0 0 " "
}

WriteWordLine 0 0 " "

#endregion TACACS Servers

#endregion Authentication Servers

#endregion authentication

#region Device Profiles
WriteWordLine 2 0 "NetScaler MAS Device Profiles"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Device Profiles"

$devprofiles = Get-vNetScalerObject -Container config -Object device_profile;
If (!$devprofiles) {
  WriteWordLine 0 0 "There are no Device Profiles configured."
}
foreach ($devprofile in $devprofiles) {

    ## IB - Use an array of hashtable to store the rows
    [System.Collections.Hashtable[]] $DEVPROFCFG = @(
    @{ Description = "Description"; Value = "Configuration"; }
    @{ Description = "Profile Name"; Value = $devprofile.name; }
    @{ Description = "Tenant ID"; Value = $devprofile.tenant_id; }
    @{ Description = "Type"; Value = GetDeviceFamily($devprofile.type); }
    @{ Description = "Use Global Setting for communivation with NetScaler"; Value = $devprofile.use_global_setting_for_communication_with_ns; }
    @{ Description = "Username"; Value = $devprofile.username; }
    @{ Description = "SNMP Community"; Value = $devprofile.snmpcommunity; }
    );

    ## IB - Create the parameters to pass to the AddWordTable function
    $Params = $null
    $Params = @{
        Hashtable = $DEVPROFCFG;
        Columns = "Description","Value";
        AutoFit = $wdAutoFitContent
        Format = -235; ## IB - Word constant for Light List Accent 5
    }
    ## IB - Add the table to the document, splatting the parameters
    $Table = AddWordTable @Params -List;

	FindWordDocumentEnd;
	$TableRange = $Null
	$Table = $Null
    WriteWordLine 0 0 " "
}

WriteWordLine 0 0 " "

#endregion Device Profiles

#endregion NetScaler MAS System Information

#region Managed Devices
WriteWordLine 1 0 "NetScaler MAS Managed Instances"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Managed Instances"

$nmasdevices = Get-vNetScalerObject -Container Config -Object managed_device

If (!$nmasdevices) {

WriteWordLine 0 0 "No NMAS Managed Devices were found."

} Else {
foreach ($nmasdevice in $nmasdevices) {
If (!$nmasdevice.hostname) { $nmasdevicename = $nmasdevice.mgmt_ip_address } Else {$nmasdevicename = $nmasdevice.hostname }

WriteWordLine 2 0 "Device: $nmasdevicename"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasdevh = @(
    @{ Description = "Device Description"; Value = $nmasdevice.description ;}
    @{ Description = "IP Address"; Value = $nmasdevice.ipv4_address ;}
    @{ Description = "Net Mask"; Value = $nmasdevice.netmask ;}
    @{ Description = "Default Gateway"; Value = $nmasdevice.gateway ;}
	@{ Description = "HA Master State"; Value = $nmasdevice.ha_master_state ;}
	@{ Description = "Instance State"; Value = $nmasdevice.instance_state ;}
    @{ Description = "Is NetScaler Gateway"; Value = $nmasdevice.gateway_deployment ;}
    @{ Description = "Model ID"; Value = $nmasdevice.model_id ;}
    @{ Description = "System Type"; Value = GetDeviceFamily($nmasdevice.type) ;}
    @{ Description = "Number of Services"; Value = $nmasdevice.sysservices ;}
    @{ Description = "Serial Number"; Value = $nmasdevice.serialnumber ;}
    @{ Description = "Profile Name"; Value = $nmasdevice.profile_name ;}
    @{ Description = "Location"; Value = $nmasdevice.location ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasdevh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      
WriteWordLine 0 0 " " 

} #end foreach

} #end if

#endregion Managed Devices
 
WriteWordLine 0 0 " "

#region Instance Groups
WriteWordLine 1 0 "NetScaler MAS Instance Groups"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Instance Groups"

$nmasdevgroups = Get-vNetScalerObject -Container Config -Object device_group

If (!$nmasdevgroups) {

WriteWordLine 0 0 "No NMAS Instance Groups were found."

} Else {
foreach ($nmasdevgroup in $nmasdevgroups) {
$nmasdevgroupname = $nmasdevgroup.name 

WriteWordLine 2 0 "Instance Group: $nmasdevgroupname"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasdevgrph = @(
    @{ Description = "Members"; Value = $nmasdevgroup.static_device_list ;}
    @{ Description = "Device Family"; Value = $nmasdevgroup.device_family ;}
    @{ Description = "Criteria Type"; Value = $nmasdevgroup.criteria_type ;}
    @{ Description = "Criteria Condition"; Value = $nmasdevgroup.criteria_condn ;}
    @{ Description = "Criteria Value"; Value = $nmasdevgroup.criteria_value ;}
    @{ Description = "Members"; Value = $nmasdevgroup.static_device_list_arr -Join ", " ;}

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasdevgrph;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "

} #end foreach

} #end if

#endregion Instance Groups
 
#region Events
WriteWordLine 1 0 "Netscaler MAS Event Management"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Event Management"

#region event rules
WriteWordLine 2 0 "NetScaler MAS Event Rules"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Event Rules"

$nmasrules = Get-vNetScalerObject -Container Config -Object event_filter

If (IsNull($nmasrules)) {

WriteWordLine 0 0 "No NMAS Event Rules were found."

} Else {
foreach ($nmasrule in $nmasrules) {
$nmasrulename = $nmasrule.name 

WriteWordLine 3 0 "Event Rule: $nmasrulename"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasruleh = @(
    @{ Description = "Enabled"; Value = $nmasrule.is_enabled ;}
    @{ Description = "Event Age Threshold (seconds)"; Value = $nmasrule.event_age_threshold ;}
    @{ Description = "Type"; Value = $nmasrule.type ;}
    @{ Description = "Severity"; Value = $nmasrule.criteria.severity ;}
    @{ Description = "Instances"; Value = $nmasrule.criteria.source -replace ",",", " ;}
    @{ Description = "Failure Objects"; Value = $nmasrule.criteria.failureobj ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasruleh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "
WriteWordLine 4 0 "Event Rule Categories: $nmasrulename"
WriteWordLine 0 0 " "

$arrCategories = $nmasrule.criteria.category.Split(",")
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $EventCatH = @();

foreach ($category in $arrCategories) {

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $EventCatH += @{
            category = $category;
        }
    }

if ($EventCatH.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $EventCatH;
        Columns = "category";
        Headers = "Category";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }
    $Table = AddWordTable @Params -List;
    FindWordDocumentEnd;
    $Table = $null
    }
}
WriteWordLine 0 0 " "

If (!$nmasrule.actions) { WriteWordLine 0 0 "No Event Actions were found." } Else {
Foreach ($ruleaction in $nmasrule.actions){
$ruleactionname = $ruleaction.profile_name

WriteWordLine 3 0 "Event Rule Action: $ruleactionname"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasruleacth = @(
    @{ Description = "Name"; Value = $ruleaction.profile_name ;}
    @{ Description = "Type"; Value = $ruleaction.type ;}
    @{ Description = "Repeat Message Notification Threshold"; Value = $ruleaction.repeat_email_notification_threshold ;}
    @{ Description = "Sender"; Value = $ruleaction.sender ;}
    @{ Description = "Subject"; Value = $ruleaction.subject ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasruleacth;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null
WriteWordLine 0 0 " "
} #end foreach action      
}
WriteWordLine 0 0 " "
} #end if
#} #end if




#endregion event rules

#region Event Settings
WriteWordLine 2 0 "Event Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Event Settings"

$nmaseventsettings = Get-vNetScalerObject -Container Config -Object trap_settings

    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasevtseth = @();

foreach ($nmaseventsetting in $nmaseventsettings) {

If (IsNull($nmaseventsetting.user_defined_severity)){$nmasevtcustom = ""} Else { $nmasevtcustom = $nmaseventsetting.user_defined_severity};


    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $nmasevtseth += @{
            Category = $nmaseventsetting.trap_category;
            DeviceFamily = GetDeviceFamily($nmaseventsetting.device_family);
            DefaultSev = $nmaseventsetting.default_severity;
            CustomSev = $nmasevtcustom;
        }
    }

if ($nmasevtseth.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $nmasevtseth;
        Columns = "Category","DeviceFamily","DefaultSev","CustomSev";
        Headers = "Category","Device Family", "Default Severity", "Custom Severity";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null
    } Else {
      WriteWordLine 0 0 " "
      WriteWordLine 0 0 "No Event Settings have been configured."
      WriteWordLine 0 0 " "
    }

#endregion Event Settings

#endregion Events

#region MAS Config Jobs
#WriteWordLine 1 0 "NetScaler MAS Configuration Jobs"
#WriteWordLine 0 0 " "

#region MAS Config Templates
WriteWordLine 1 0 "NetScaler MAS Configuration Templates"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Configuration Templates"

$nmasconfs = Get-vNetScalerObject -Container Config -Object configuration_template

If (!$nmasconfs) {

WriteWordLine 0 0 "No NMAS Configuration Templates were found."

} Else {
foreach ($nmasconf in $nmasconfs) {
$nmasconfname = $nmasconf.name 

WriteWordLine 2 0 "Configuration Template: $nmasconfname"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasconfh = @(
    @{ Description = "Device Family"; Value = GetDeviceFamily($nmasconf.device_family) ;}
    @{ Description = "Variables"; Value = $nmasconf.variables -Join ", " ;}
    @{ Description = "Description"; Value = $nmasconf.description ;}
    @{ Description = "Built-in Template"; Value = $nmasconf.is_inbuilt ;}
    @{ Description = "Visible"; Value = $nmasconf.is_visible ;}

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasconfh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null     
WriteWordLine 0 0 " " 
If ($nmasconf.commands) {
WriteWordLine 3 0 "Commands"
WriteWordLine 0 0 " "
[System.Collections.Hashtable[]] $nmascommandh = @();
  ForEach ($nmasconfcommand in $nmasconf.commands) {

    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $nmascommandh += @{
            Protocol = $nmasconfcommand.protocol;
            Command = $nmasconfcommand.command;
            
        }
    }
    $Params = $null
    $Params = @{
        Hashtable = $nmascommandh;
        Columns = "Protocol","Command";
        Headers = "Protocol","Command";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null
    
  }
}
WriteWordLine 0 0 " "



} #end if
#endregion MAS Config Templates

#endregion MAS Config Jobs

#region MAS Datacenters
WriteWordLine 1 0 "NetScaler MAS Datacenters"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Datacenters"

$nmasdcs = Get-vNetScalerObject -Container Config -Object mps_datacenter

    
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasdch = @();

foreach ($nmasdc in $nmasdcs) {


    ## IB - Create parameters for the hashtable so that we can splat them otherwise the
    ## IB - command will be about 400 characters wide!
    $nmasdch += @{
            DCName = $nmasdc.name;
            IPBlocks = $nmasdc.ip_block_array -Join ", ";
        }
    }


    $Params = $null
    $Params = @{
        Hashtable = $nmasdch;
        Columns = "DCName","IPBlocks";
        Headers = "Datacenter Name", "IP Blocks";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    $Table = $null

#region MAS IP Blocks
WriteWordLine 2 0 "NetScaler MAS IP Blocks"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS IP Blocks"

$nmasipbs = Get-vNetScalerObject -Container Config -Object ip_block

If (!$nmasipbs) {

WriteWordLine 0 0 "No NMAS IP Blocks have been defined."

} Else {
foreach ($nmasipb in $nmasipbs) {
$nmasipbname = $nmasipb.name 

WriteWordLine 3 0 "IP Block: $nmasipbname"
WriteWordLine 0 0 " "
## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasipbh = @(
    @{ Description = "Start IP"; Value = $nmasipb.start_ip ;}
    @{ Description = "End IP"; Value = $nmasipb.end_ip ;}
    @{ Description = "Country"; Value = $nmasipb.country ;}
    @{ Description = "Region"; Value = $nmasipb.region ;}
    @{ Description = "City"; Value = $nmasipb.City ;}
    @{ Description = "Region Code"; Value = $nmasipb.region_code ;}
    @{ Description = "Country Code"; Value = $nmasipb.country_code ;}
    @{ Description = "Latitude"; Value = $nmasipb.latitude ;}
    @{ Description = "Longitude"; Value = $nmasipb.longitude ;}
    @{ Description = "Custom City"; Value = $nmasipb.custom_city ;}
    @{ Description = "Custom Country"; Value = $nmasipb.custom_country ;}
    @{ Description = "Custom Region"; Value = $nmasipb.custom_region ;}
    

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasipbh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "

} #end foreach

} #end if


#endregion MAS IP Blocks

#endregion MAS Datacenters

#region MAS Stylebooks
WriteWordLine 1 0 "NetScaler MAS StyleBooks"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS StyleBooks"

$Stylebooks = Get-vNetScalerStylebooks

foreach ($Stylebook in $Stylebooks) {

$StylebookName = $Stylebook.display_name

WriteWordLine 2 0 "StyleBook: $StylebookName"
WriteWordLine 0 0 " "
$StylebookInfo = Get-vNetScalerStylebookObject -NameSpace $Stylebook.namespace -Version $Stylebook.version -Name $Stylebook.name;


[System.Collections.Hashtable[]] $nmassbh = @(
    @{ Description = "Description"; Value = $Stylebook.description -replace '(?:\r|\n)', '' ;}
    @{ Description = "Version"; Value = $Stylebook.version ;}
    @{ Description = "Date Imported"; Value = $Stylebook.imported_date ;}
    @{ Description = "NameSpace"; Value = $Stylebook.namespace ;}
    $ImportedStyleBooks = $Stylebook.uses_stylebooks.Name;
    @{ Description = "Imported Stylebooks"; Value = $ImportedStyleBooks -Join ", " ;}    

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmassbh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "

WriteWordLine 3 0 "StyleBook Source"
WriteWordLine 0 0 " "
$SBSource = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($StylebookInfo.source))
WriteWordLine 0 0 $SBSource
WriteWordLine 0 0 " "


} #end foreach Stylebook



#endregion MAS Stylebooks

#region MAS Analytics Settings

WriteWordLine 1 0 "NetScaler MAS Analytics Settings"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Analytics Settings"


$nmasreporttz = Get-vNetScalerObject -Container Config -Object report_timezone;
$nmasicatimeout = Get-vNetScalerObject -Container Config -Object ica_session_timeout;
$nmasmultihop = Get-vNetScalerObject -Container Config -Object multihop_feature;
$nmasadaptive = Get-vNetScalerObject -Container Config -Object adaptive_threshold_feature;
$nmasdbindex = Get-vNetScalerObject -Container Config -Object af_database_index;
$nmasdbcleanup = Get-vNetScalerObject -Container Config -Object af_database_cleanup;
$nmasdbcache = Get-vNetScalerObject -Container Config -Object db_cache;
$nmaslogdata = Get-vNetScalerObject -Container Config -Object log_datarecord;
$nmasdataduration = Get-vNetScalerObject -Container Config -Object af_database_duration;
$nmashttpheader = Get-vNetScalerObject -Container Config -Object http_header;
$nmasslacollection = Get-vNetScalerObject -Container Config -Object sla_collection_enable;
$nmasurlcollection = Get-vNetScalerObject -Container Config -Object af_url_collection;
$nmasurlparameter = Get-vNetScalerObject -Container Config -Object url_parameter;

## IB - Use an array of hashtable to store the rows
[System.Collections.Hashtable[]] $nmasanalyticsh = @(
    @{ Description = "Dashboard Reporting Time Zone Setting"; Value = $nmasreporttz.timezone ;}
	@{ Description = "ICA Session Timeout Value (minutes)"; Value = $nmasicatimeout.timeout_value ;}
	@{ Description = "Enable Mutihop"; Value = $nmasmultihop.enable ;}
    @{ Description = "Enable Adaptive Threshold"; Value = $nmasadaptive.enable ;}
    @{ Description = "Enable Database Indexing"; Value = $nmasdbindex.enable ;}
    @{ Description = "Enable Database Cleanup"; Value = $nmasdbcleanup.enable ;}
    @{ Description = "Enable Database Cache"; Value = $nmasdbcache.enable ;}
    @{ Description = "Enable HDX Insight Logs"; Value = $nmaslogdata.hdx ;}
    @{ Description = "Enable Web Insight Logs"; Value = $nmaslogdata.web ;}
    @{ Description = "Enable CB WAN Insight Logs"; Value = $nmaslogdata.cbwan ;}
    @{ Description = "Enable Security Insight Logs"; Value = $nmaslogdata.security ;}
    @{ Description = "Duration for data to persist (days)"; Value = $nmasdataduration.days ;}
    @{ Description = "Show HTTP Request Method Report"; Value = $nmashttpheader.http_req_method ;}
    @{ Description = "Show HTTP Response Status Report"; Value = $nmashttpheader.http_resp_status ;}
    @{ Description = "Show User Agent Report"; Value = $nmashttpheader.user_agent ;}
    @{ Description = "Show Operating System Report"; Value = $nmashttpheader.operating_system ;}
    @{ Description = "Show Domain Report"; Value = $nmashttpheader.domain ;}
    @{ Description = "Show Content Type Report"; Value = $nmashttpheader.http_content_type ;}
    @{ Description = "Show Media Type Report"; Value = $nmashttpheader.http_media_type ;}
    @{ Description = "Show Uncached Resources Report"; Value = $nmashttpheader.ic_nostore_reason ;}
    @{ Description = "Enable SLA Data Collection"; Value = $nmasslacollection.sla_enable ;}
    @{ Description = "Enable URL Data Collection"; Value = $nmasurlcollection.enable ;}
    @{ Description = "Trim URL Parameters"; Value = $nmasurlparameter.remove ;}

);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasanalyticsh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null  

WriteWordLine 0 0 " "

#region DB Summarization
WriteWordLine 2 0 "NetScaler MAS Database Summarization"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Database Summarization"

$nmasdbsummrules = Get-vNetScalerObject -Container Config -Object db_summarization_config;

[System.Collections.Hashtable[]] $nmasdbsummh = @();

foreach($nmasdbsummrule in $nmasdbsummrules) {

$minutelyspan = New-Timespan -Seconds $nmasdbsummrule.minute
$hourlyspan = New-Timespan -Seconds $nmasdbsummrule.hourly
$dailyspan = New-Timespan -Seconds $nmasdbsummrule.daily
$minutely = $minutelyspan.Hours
$hourly = $hourlyspan.Days
$daily = $dailyspan.Days

$nmasdbsummh += @{
            Name = $nmasdbsummrule.name;
            Minutely = "$minutely Hours";
            Hourly = "$hourly Days";
            Daily = "$daily Days";
        }



} #end foreach

    $Params = $null
    $Params = @{
        Hashtable = $nmasdbsummh;
        Columns = "Name","Minutely","Hourly","Daily";
        Headers = "Insight Name", "Hours to persiste minutely data", "Days to persist hourly data","Days to persist daily data";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "


#endregion DB Summarization

#region Adaptive Thresholds
WriteWordLine 2 0 "NetScaler MAS Adaptive Thresholds"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Analytics Adaptive Thresholds"

$nmasadaptivethresholds = Get-vNetScalerObject -Container Config -Object adaptive_threshold;

[System.Collections.Hashtable[]] $nmasadapth = @();

foreach($nmasadaptivethreshold in $nmasadaptivethresholds) {


$nmasadapth += @{
            Name = $nmasadaptivethreshold.threshold_name;
            Value = $nmasadaptivethreshold.threshold_value;
            Duration = $nmasadaptivethreshold.monitor_duration;
            Type = $nmasadaptivethreshold.resource_type;
        }



} #end foreach

If ($nmasadapth.Length -gt 0) {
    $Params = $null
    $Params = @{
        Hashtable = $nmasadapth;
        Columns = "Name","Value","Duration","Type";
        Headers = "Name", "Threshold Multiplier", "Duration","Entity";
        Format = -235; ## IB - Word constant for Light Grid Accent 5 (could use -207 for Accent 3 (grey))
        AutoFit = $wdAutoFitContent;
        }

    $Table = AddWordTable @Params;
    FindWordDocumentEnd;
    WriteWordLine 0 0 " "
    } Else {
    WriteWordLine 0 0 " "
    WriteWordLine 0 0 "No adaptive thresholds have been configured."
    WriteWordLine 0 0 " "
    }


#endregion Adaptive Thresholds

#region thresholds and alerts
WriteWordLine 2 0 "NetScaler MAS Thresholds"
WriteWordLine 0 0 " "

$Chapter++
Write-Verbose "$(Get-Date): Chapter $Chapter/$Chapters NetScaler MAS Analytics Thresholds"

$nmasthresholds = Get-vNetScalerObject -Container Config -Object insight_threshold;

If (IsNull($nmasthresholds)) {
    WriteWordLine 0 0 " "
    WriteWordLine 0 0 "No thresholds have been configured."
    WriteWordLine 0 0 " "
} Else {

  Foreach($nmasthreshold in $nmasthresholds) {

  [System.Collections.Hashtable[]] $nmasthresholdh = @(
    @{ Description = "Name"; Value = $nmasthreshold.name ;}
    @{ Description = "Alerts Enabled"; Value = $nmasthreshold.is_enabled ;}
    @{ Description = "Traffic Type"; Value = $nmasthreshold.group_name ;}
    @{ Description = "Entity"; Value = $nmasthreshold.resource_type ;}
    @{ Description = "Reference Key"; Value = $nmasthreshold.reference_key ;}
    @{ Description = "Rule"; Value = $nmasthreshold.rule ;}
    @{ Description = "Duration"; Value = GetDuration($nmasthreshold.duration) ;}
);

## IB - Create the parameters to pass to the AddWordTable function
$Params = $null
$Params = @{
    Hashtable = $nmasthresholdh;
    Columns = "Description","Value";
    AutoFit = $wdAutoFitContent
    Format = -235; ## IB - Word constant for Light List Accent 5
}
## IB - Add the table to the document, splatting the parameters
$Table = AddWordTable @Params;

FindWordDocumentEnd;
$TableRange = $Null
$Table = $Null      

WriteWordLine 0 0 " "

  } #end foreach

}

#endregion thresholds and alerts


#endregion MAS Analytics Settings



#endregion NetScaler MAS Documentation Script Complete

#region script template 2

Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

###Change the two lines below for your script
$AbstractTitle = "NetScaler MAS Documentation Report"
$SubjectTitle = "NetScaler MAS Documentation Report"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#recommended by webster
#$error
#endregion script template 2