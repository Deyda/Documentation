#Requires -Version 3.0
#requires -Module ActiveDirectory
#requires -Module GroupPolicy


<#
.SYNOPSIS
	Creates a Backup and Reports for all Group Policies in the current Active Directory domain.
.DESCRIPTION
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
.PARAMETER ADDomain
	Specifies an Active Directory domain object by providing one of the following 
	property values. The identifier in parentheses is the LDAP display name for the 
	attribute. All values are for the domainDNS object that represents the domain.

	Distinguished Name

	Example: DC=tullahoma,DC=corp,DC=labaddomain,DC=com

	GUID (objectGUID)

	Example: b9fa5fbd-4334-4a98-85f1-3a3a44069fc6

	Security Identifier (objectSid)

	Example: S-1-5-21-3643273344-1505409314-3732760578

	DNS domain name

	Example: tullahoma.corp.labaddomain.com

	NetBIOS domain name

	Example: Tullahoma

	Default value is $Env:USERDNSDOMAIN

	If both ADDomain and OrganizationalUnit are used, the latter takes preference.
.PARAMETER ComputerName
	Specifies which domain controller to use to run the script against.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	This parameter has an alias of ServerName.
	Default value is $Env:USERDNSDOMAIN	
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
	
	The folder specified must already exist.
.PARAMETER GPOFilter
	Specifies a text string to restrict GPOs retrieve.
	
	If GPOFilter is XenApp, then every GPO in the ADDomain or OrganizationalUnit
	That has "XenApp" anywhere in the name is backed up and has reports
	generated.
	
	Default is all GPOs.
.PARAMETER MaxZipSize
	Specifies the maximum size, in MB, of the two Zip files created.
	Default is 150MB which is the limit of Outlook 365 email attachments.
	
	https://technet.microsoft.com/en-us/library/exchange-online-limits.aspx#MessageLimits
.PARAMETER OrganizationalUnit
	Restricts the retrieval of computer accounts to a specific OU tree. 
	Must be entered in Distinguished Name format. i.e. OU=XenDesktop,DC=domain,DC=tld. 
	
	The script GPOs from the top level OU and all sub-level OUs.
	
	If both ADDomain and OrganizationalUnit are used, the latter takes preference.
	
	Alias OU
.PARAMETER Rename
	If a GPO name contains any the following characters, <>:"\/|?*, replace the character 
	with an _ (Underscore), allowing the creation of a report file in Windows.
	
	The GPO is then renamed in the Group Policy Management Console and Active Directory.

	The default is False.
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
	PS C:\PSScript > .\Docu-GPO.ps1
	
	ComputerName = $Env:USERDNSDOMAIN
	ADDomain = $Env:USERDNSDOMAIN
	Folder = $pwd
.EXAMPLE
	PS C:\PSScript > .\Docu-GPO.ps1 -ComputerName PDCeDC
	
	ComputerName = PDCeDC
	ADDomain = $Env:USERDNSDOMAIN
	Folder = $pwd
.EXAMPLE
	PS C:\PSScript > .\Docu-GPO.ps1 -ComputerName ChildPDCeDC 
	-ADDomain ChildDomain.com
	
	Assuming the script is run from the parent domain.
	ComputerName = ChildPDCeDC
	ADDomain = ChildDomain.com
	Folder = $pwd
.EXAMPLE
	PS C:\PSScript > .\Docu-GPO.ps1 -ComputerName ChildPDCeDC 
	-ADDomain ChildDomain.com -Folder c:\GPOReports
	
	Assuming the script is run from the parent domain.
	ComputerName = ChildPDCeDC
	ADDomain = ChildDomain.com
	Folder = C:\GPOReports (C:\GPOReports must already exist)
.EXAMPLE
	PS C:\PSScript > .\Docu-GPO.ps1 -SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld -To ITGroup@domain.tld	
	
	The script will use the email server mail.domain.tld, sending from 
	XDAdmin@domain.tld, sending to ITGroup@domain.tld.
	
	The script will use the default SMTP port 25 and will not use SSL.
	
	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\Docu-GPO.ps1 -SmtpServer smtp.office365.com 
	-SmtpPort 587 -UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	
	
	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.
	
	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS c:\PSScript > .\Docu-GPO.ps1 -folder c:\gpobackups 
	-ou "ou=lab,dc=labaddomain,dc=com" -gpofilter "xenapp"
	
	Processes all GPOs in the OU "lab" in the domain LabAdomain.com 
	that contain "xenapp" anywhere in the GPO name.
.EXAMPLE
	PS c:\PSScript > .\Docu-GPO.ps1 -folder c:\gpobackups 
	-ou "ou=lab,dc=labaddomain,dc=com"
	
	Processes all GPOs in the OU "lab" in the domain LabAdomain.com.
.EXAMPLE
	PS c:\PSScript > .\Docu-GPO.ps1 -folder c:\gpobackups 
	-ou "ou=lab,dc=labaddomain,dc=com" -GPOFilter ""
	
	Processes all GPOs in the OU "lab" in the domain LabADDomain.com.
.EXAMPLE
	PS c:\PSScript > .\Docu-GPO.ps1 -folder c:\gpobackups 
	-ou ""
	
	Processes all GPOs in all OUs in the default domain of $Env:USERDNSDOMAIN.
.EXAMPLE
	PS C:\PSScript > .\Docu-GPO.ps1 -Rename
	
	ComputerName = $Env:USERDNSDOMAIN
	ADDomain = $Env:USERDNSDOMAIN
	Folder = $pwd
	
	Replace Windows invalid filename charcters in the GPO name with an _ (Underscore).
	The GPO is then renamed in the Group Policy Management Console and Active Directory. 
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.
.NOTES
	NAME: Docu-GPO.ps1
	VERSION: 1.23
	AUTHOR: Carl Webster
	LASTEDIT: September 12, 2019
#>


[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Default") ]

Param(
	[parameter(Mandatory=$False)] 
	[string]$ADDomain=$Env:USERDNSDOMAIN, 

	[parameter(Mandatory=$False)] 
	[Alias("ServerName")]
	[string]$ComputerName=$Env:USERDNSDOMAIN,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[string]$GPOFilter="",
	
	[parameter(Mandatory=$False)] 
	[int]$MaxZipSize=150,
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[Alias("OU")]
	[string]$OrganizationalUnit="",
	
	[parameter(ParameterSetName="Default",Mandatory=$False)] 
	[switch]$Rename=$False,

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
#Created on April 25, 2018

#Version 1.23 12-Sep-2019
#	Add a Rename parameter switch. (Thanks to Jani Kohonen for this suggestion)
#	If a GPO name contains any the following characters, <>:"\/|?*, replace the character 
#	with an _ (Underscore), allowing the creation of a report file in Windows.
#	The GPO is then renamed in the Group Policy Management Console and Active Directory.
#	The default is False.

#Version 1.22 24-Apr-2019
#	Fixed bug when a GPO had been linked multiple times. The script counted all the GPOs, 
#		even the duplicates. Added -Unique to the Sort.
#		If a GPO was linked ten times, the script backed up the GPO ten times and created
#		the reports ten times.
#
#Version 1.21 11-Apr-2019
#	If -OrganizationUnit parameter was used, the $ADDomain paramater was blanked out and
#		not available for the Zip file creation. IF $ADDomain is blank, set $ADDomain to
#		$Env:USERDNSDOMAIN
#
#Version 1.20 29-Mar-2019
#	Received performance tuning help from Michael B. Smith for this update
#	Add requirement for the ActiveDirectory module
#	Add simple error checking to the backup and report creation
#	Add two parameters: GPOFilter and OrganizationalUnit
#	Adding testing for illegal filename characters '<>:"\/|?*' in the GPO name before 
#	creating the HTML and XML report files.
#	If both ADDomain and OrganizationalUnit are specified, set ADDomain to empty string
#	Updated Function ProcessScriptEnd with new parameters
#	Updated help text
#
#Version 1.10 1-Jun-2018
#
#	Add Parameter MaxZipSize (in MBs)
#		Test combined size of the two Zip files created to see if <= to MaxZipSize
#		If > MaxZipSize, do not attempt the email
#	Backup GPOs one at a time
#		Show the name of each GPO being backed up to show the GPO causing a backup error
#
#Version 1.0 released to the community on 1-May-2018
#


Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

#If $OrganizationalUnit is used, blank out $ADDomain
If(![String]::IsNullOrEmpty($ADDomain) -and ![String]::IsNullOrEmpty($OrganizationalUnit)) #V1.20
{
	#OrganizationalUnit was specified, blank out ADDomain
	$ADDomain = ""
}

#region check for DA and elevatation
Function UserIsaDomainAdmin
{
	#function adapted from sample code provided by Thomas Vuylsteke
	$IsDA = $False
	$name = $env:username
	If([String]::IsNullOrEmpty($ADDomain)) 
	{
		$ADDomain = $Env:USERDNSDOMAIN
	}
	
	Write-Verbose "$(Get-Date): TokenGroups - Checking groups for $name"

	$root = [ADSI]""
	$filter = "(sAMAccountName=$name)"
	$props = @("distinguishedName")
	$Searcher = new-Object System.DirectoryServices.DirectorySearcher($root,$filter,$props)
	$account = $Searcher.FindOne().properties.distinguishedname

	$user = [ADSI]"LDAP://$Account"
	$user.GetInfoEx(@("tokengroups"),0)
	$groups = $user.Get("tokengroups")

	$domainAdminsSID = New-Object System.Security.Principal.SecurityIdentifier (((Get-ADDomain -Server $ADDomain -EA 0).DomainSid).Value+"-512") 

	ForEach($group in $groups)
	{     
		$ID = New-Object System.Security.Principal.SecurityIdentifier($group,0)       
		If($ID.CompareTo($domainAdminsSID) -eq 0)
		{
			$IsDA = $True
			Break
		}     
	}

	Return $IsDA
}

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
		Write-Host "" -Foreground White
		Write-Host "$(Get-Date): This is NOT an elevated PowerShell session" -Foreground White
		Write-Host "" -Foreground White
		Return $False
	}
}
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$Script:GPOStartTime = Get-Date

	#make sure user is running the script with Domain Admin rights
	Write-Verbose "$(Get-Date): Testing to see if $env:username has Domain Admin rights"
	$AmIReallyDA = UserIsaDomainAdmin
	If($AmIReallyDA -eq $True)
	{
		#user has Domain Admin rights
		Write-Verbose "$(Get-Date): $env:username has Domain Admin rights in the $ADDomain Domain"
	}
	Else
	{
		#user does not have Domain Admin rights
		Write-Error "$(Get-Date): $env:username does not have Domain Admin rights in the $ADDomain Domain. Script cannot continue."
		$ErrorActionPreference = $SaveEAPreference
		Exit
	}
	
	$Elevated = ElevatedSession
	
	If( -not $Elevated )
	{
		Write-Error "This is not an elevated PowerShell session. Script cannot continue."
		$ErrorActionPreference = $SaveEAPreference
		Exit
	}

	#if computer name is localhost, get actual server name
	If($ComputerName -eq "localhost")
	{
		$ComputerName = $env:ComputerName
		Write-Verbose "$(Get-Date): Server name has been changed from localhost to $($ComputerName)"
	}
	
	#see if default value of $Env:USERDNSDOMAIN was used
	If($ComputerName -eq $Env:USERDNSDOMAIN)
	{
		#change $ComputerName to a found global catalog server
		$Results = (Get-ADDomainController -DomainName $ADDomain -Discover -Service GlobalCatalog -EA 0).Name
		
		If($? -and $Null -ne $Results)
		{
			$ComputerName = $Results
			Write-Verbose "$(Get-Date): Server name has been changed from $Env:USERDNSDOMAIN to $ComputerName"
		}
		ElseIf(!$?) #changed for 2.16
		{
			#may be in a child domain where -Service GlobalCatalog doesn't work. Try PrimaryDC
			$Results = (Get-ADDomainController -DomainName $ADDomain -Discover -Service PrimaryDC -EA 0).Name

			If($? -and $Null -ne $Results)
			{
				$ComputerName = $Results
				Write-Verbose "$(Get-Date): Server name has been changed from $Env:USERDNSDOMAIN to $ComputerName"
			}
		}
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $ComputerName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$ComputerName = $Result.HostName
			Write-Verbose "$(Get-Date): Server name has been updated from $ip to $ComputerName"
		}
		Else
		{
			Write-Warning "Unable to resolve $ComputerName to a hostname"
		}
	}
	Else
	{
		#server is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}

	If(![String]::IsNullOrEmpty($ComputerName)) 
	{
		#get server name
		#first test to make sure the server is reachable
		Write-Verbose "$(Get-Date): Testing to see if $ComputerName is online and reachable"
		If(Test-Connection -ComputerName $ComputerName -quiet -EA 0)
		{
			Write-Verbose "$(Get-Date): Server $ComputerName is online."
			Write-Verbose "$(Get-Date): `tTest #1 to see if $ComputerName is a Domain Controller."
			#the server may be online but is it really a domain controller?

			#is the ComputerName in the current domain
			$Results = Get-ADDomainController $ComputerName -EA 0
			
			If(!$? -or $Null -eq $Results)
			{
				#try using the Forest name
				Write-Verbose "$(Get-Date): `tTest #2 to see if $ComputerName is a Domain Controller."
				$Results = Get-ADDomainController $ComputerName -Server $ADDomain -EA 0
				If(!$?)
				{
					$ErrorActionPreference = $SaveEAPreference
					Write-Error "`n`n`t`t$ComputerName is not a domain controller for $ADDomain.`n`t`tScript cannot continue.`n`n"
					$ErrorActionPreference = $SaveEAPreference
					Exit
				}
				Else
				{
					Write-Verbose "$(Get-Date): `tTest #2 succeeded. $ComputerName is a Domain Controller."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): `tTest #1 succeeded. $ComputerName is a Domain Controller."
			}
			
			$Results = $Null
		}
		Else
		{
			Write-Error "Computer $ComputerName is offline.`nScript cannot continue.`n`n"
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
	}

	If(![String]::IsNullOrEmpty($ADDomain))
	{
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$results = Get-ADDomain -Identity $ADDomain -EA 0
			
			If(!$?)
			{
				Write-Error "Could not find a domain identified by: $ADDomain.`nScript cannot continue.`n`n"
				$ErrorActionPreference = $SaveEAPreference
				Exit
			}
		}
		Else
		{
			$results = Get-ADDomain -Identity $ADDomain -Server $ComputerName -EA 0

			If(!$?)
			{
				Write-Error "Could not find a domain with the name of $ADDomain.`n`n`t`tScript cannot continue.`n`n`t`tIs $ComputerName running Active Directory Web Services?"
				$ErrorActionPreference = $SaveEAPreference
				Exit
			}
		}
		Write-Verbose "$(Get-Date): $ADDomain is a valid domain name"
	}
	ElseIf(![String]::IsNullOrEmpty($OrganizationalUnit))
	{
		Write-Verbose "$(Get-Date): Validating Organnization Unit"
		try 
		{
			$results = Get-ADOrganizationalUnit -Identity $OrganizationalUnit
		} 
		
		catch
		{
			#does not exist
			Write-Error "Organization Unit $OrganizationalUnit does not exist.`n`nScript cannot continue`n`n"
			Exit
		}	
		
		Write-Verbose "$(Get-Date): $OrganizationalUnit is a valid OU"
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
				$ErrorActionPreference = $SaveEAPreference
				Exit
			}
		}
		Else
		{
			#does not exist
			Write-Error "Folder $Folder does not exist.  Script cannot continue"
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
	}

	If($Folder -eq "")
	{
		$Script:pwdPath = $pwd.Path
	}
	Else
	{
		$Script:pwdPath = $Folder
	}

	If($Script:pwdPath.EndsWith("\"))
	{
		#remove the trailing \
		$Script:pwdPath = $Script:pwdPath.SubString(0, ($Script:pwdPath.Length - 1))
	}

	If($Dev)
	{
		$Error.Clear()
		$Script:DevErrorFile = "$($Script:pwdPath)\Get-GPOBackupAndReportsScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	If($Log) 
	{
		#start transcript logging
		$Script:LogPath = "$($Script:pwdPath)\Get-GPOBackupAndReportsScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		
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

	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

	#enter the email creds
	If(![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$Script:emailCredentials = Get-Credential -Message "Enter the email account and password to send email"
	}
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"
	
	$Success = $True
	$emailAttachment = $Attachments
	$emailSubject = "GPO Backup and Reports for $ADDomain"
	$emailBody = @"
Hello, <br />
<br />
$emailsubject is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()

	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL -credential $emailCredentials *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To -credential $emailCredentials *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

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

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
			$Success = $False
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
		$Success = $False
	}

    If($Success)
    {
        Return $True
    }
    Else
    {
        Return $False
    }
}
#endregion

#region main function
Function GetGpoBackupAndReports
{
	If(![String]::IsNullOrEmpty($OrganizationalUnit))
	{
		Write-Verbose "$(Get-Date): Retrieving all GPOs in OU $OrganizationalUnit"
		$GPOs = New-Object System.Collections.ArrayList
		$LinkedGPOs = Get-ADOrganizationalUnit -Filter * -SearchBase $OrganizationalUnit -SearchScope Subtree  | Select-Object -ExpandProperty LinkedGroupPolicyObjects
		
		If($? -and $Null -ne $LinkedGPOs)
		{
			ForEach($LinkedGPO in $LinkedGPOs) 
			{             
				$GpoName = ([adsi]"LDAP://$LinkedGPO").DisplayName.ToString()
				If($GpoName -like "*$GpoFilter*")
				{
					$GPOs.Add($GPOName) > $Null
				}
			}
		}
		ElseIf($? -and $Null -eq $LinkedGPOs)
		{
			Write-Warning "No GPOs were found. Script cannot continue."
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
		Else
		{
			Write-Error "Unable to retrieve GPOs. Script cannot continue."
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
	}
	ElseIf(![String]::IsNullOrEmpty($ADDomain))
	{
		Write-Verbose "$(Get-Date): Retrieving all GPOs in $ADDomain"
		$GPOs = New-Object System.Collections.ArrayList
		
		$results = @(Get-GPO -All -EA 0)
		
		If($? -and $Null -ne $results)
		{
			$results | Where-Object { $_.DisplayName -like "*$GpoFilter*" } | ForEach-Object { $null = $GPOs.Add( $_.DisplayName ) }
		}
		ElseIf($? -and $Null -eq $results)
		{
			Write-Warning "No GPOs were found. Script cannot continue."
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
		Else
		{
			Write-Error "Unable to retrieve GPOs. Script cannot continue."
			$ErrorActionPreference = $SaveEAPreference
			Exit
		}
	}

	If($Null -ne $GPOs)
	{
		$GPOs = $GPOs | Sort-Object -Unique
		
		[int]$GPOCnt = $GPOs.Count

		Write-Verbose "$(Get-Date): Creating Backup and Reports folders"
		$ScriptDir = $Script:pwdPath
		$BackupDir = "$($ScriptDir)\GPOBackups"
		$ReportsDir = "$($ScriptDir)\GPOReports"
		New-Item -Path $BackupDir -Force -ItemType "directory" *> $Null
		New-Item -Path $ReportsDir -Force -ItemType "directory" *> $Null

		Write-Verbose "$(Get-Date): Backing up $GPOCnt GPOs to $BackupDir"
		
		#V1.10, backup individual GPO to show which GPO causes the backup to fail
		
		ForEach($GPO in $GPOs)
		{
			Write-Verbose "$(Get-Date): Backing up GPO $GPO"

			$Null = Backup-GPO -Name $GPO -Path $BackupDir -EA 0 *> $Null
				
			If(!($?))
			{
				Write-Warning "`t`t`tUnable to backup GPO $($GPO): $error[0].exception.ToString()"
			}
		}

		ForEach($GPO in $GPOs)
		{
			Write-Verbose "$(Get-Date): Creating reports for GPO $GPO"
			
			#Version 1.20, add checking the GPO name for illegal filename characters.
			#for GPO backups, the GUID is used, for GPO reports, the GPO name is used.
			#you can use characters in a GPO name that are illegal for Windows filenames.
			#Using any the following characters '<>:"\/|?*' in a GPO name means Windows 
			#can't create an HTML or XML file for the GPO report.
			
			$Skip = $False #V1.23
			If($GPO.IndexOfAny( '<>:"\/|?*' ) -ge 0)
			{
				If($Rename -eq $False) #default behaviour
				{
					$Skip = $True #V1.23
					Write-Host "WARNING: GPO `"$GPO`" contains illegal filename characters. No report created." -Foreground White 
				}
				ElseIf($Rename -eq $True) #V1.23 new option
				{
					Write-Host "GPO `"$GPO`" contains illegal filename characters. Replacing with _" -Foreground White 
					#replace each of the invalid characters with an _ (underscore)
					$NewGPOName = $GPO
					$NewGPOName = $NewGPOName.Replace('<','_')
					$NewGPOName = $NewGPOName.Replace('>','_')
					$NewGPOName = $NewGPOName.Replace(':','_')
					$NewGPOName = $NewGPOName.Replace('"','_')
					$NewGPOName = $NewGPOName.Replace('\','_')
					$NewGPOName = $NewGPOName.Replace('/','_')
					$NewGPOName = $NewGPOName.Replace('|','_')
					$NewGPOName = $NewGPOName.Replace('?','_')
					$NewGPOName = $NewGPOName.Replace('*','_')
					Write-Host "New GPO name is `"$NewGPOName`"" -Foreground White 
					
					$results = Rename-GPO -Name $GPO -TargetName $NewGPOName
					
					If($? -and $Null -ne $results)
					{
						Write-Host ""
						Write-Host "GPO successfully renamed" -Foreground White 
						Write-Host ""
						$GPO = $NewGPOName
					}
				}
			}
			
			If($Skip -eq $False) #V1.23
			{
				try
				{
					$Results = Get-GPOReport -Name $GPO -ReportType HTML -Path "$($ReportsDir)\$($GPO).html"
				}
				
				catch
				{
					Write-Warning "Unable to create an HTML report for GPO $GPO : $error[0].Exception.ToString()"
				}

				try
				{
					$Results = Get-GPOReport -Name $GPO -ReportType XML -Path "$($ReportsDir)\$($GPO).xml"
				}
				
				catch
				{
					Write-Warning "Unable to create an XML report for GPO $GPO : $error[0].exception.ToString()"
				}
			}
		}
	}
	ElseIf($Null -eq $GPOs)
	{
		Write-Warning "No GPOs were found. Script cannot continue."
		$ErrorActionPreference = $SaveEAPreference
		Exit
	}
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	Write-Verbose "$(Get-Date): Creating First Zip File"
	$ScriptDir = $Script:pwdPath
	$BackupDir = "$($ScriptDir)\GPOBackups"
	$ReportsDir = "$($ScriptDir)\GPOReports"
	If([String]::IsNullOrEmpty($ADDomain)) #added V1.21
	{
		$ADDomain = $Env:USERDNSDOMAIN
	}

	[Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" ) > $Null
	[System.AppDomain]::CurrentDomain.GetAssemblies() > $Null
	$src_folder = $BackupDir
	$destfile1 = "$ScriptDir\_GetGPOBackup_$($ADDomain).zip"
	$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
	$includebasedir = $true
	[System.IO.Compression.ZipFile]::CreateFromDirectory($src_folder, $destfile1, $compressionLevel, $includebasedir) > $Null

	Write-Verbose "$(Get-Date): Creating second Zip File"
	$src_folder = $ReportsDir
	$destfile2 = "$ScriptDir\_GetGPOReports_$($ADDomain).zip"
	$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
	$includebasedir = $true
	[System.IO.Compression.ZipFile]::CreateFromDirectory($src_folder, $destfile2, $compressionLevel, $includebasedir) > $Null

	#email zip files if requested
	If(![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		#V1.10, check the size of the combined zip files to see if the size is smaller than MaxZipSize (150MB by default)
		$Zip1Size = ([System.IO.FileInfo]$destfile1).Length
		$Zip1Size = ([math]::round(($Zip1Size/1MB),3))

		$Zip2Size = ([System.IO.FileInfo]$destfile2).Length
		$Zip2Size = ([math]::round(($Zip2Size/1MB),3))
		
		$TotalZipSize = $Zip1Size + $Zip2Size

		If($TotalZipSize -le $MaxZipSize)
		{
			Write-Verbose "$(Get-Date): Combined size of the two zip files is $TotalZipSize MB. Proceeding with email attempt."
			Write-Verbose "$(Get-Date): Sending email"
			$emailattachment = @()
			$emailAttachment += $destfile1
			$emailAttachment += $destfile2
			$MailSuccess = SendEmail $emailAttachment
			If($MailSuccess)
			{
				Write-Verbose "$(Get-Date): Email sent"
			}
			Else
			{
				Write-Host "$(Get-Date): Unable to send $emailAttachment via email" -ForegroundColor Red
				Write-Host "$(Get-Date): Please send $emailAttachment to $To" -ForegroundColor Red
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date): Combined size of the two zip files is $TotalZipSize. Cancel email attempt."
			Write-Host "$(Get-Date): Unable to send $emailAttachment via email because of MaxZipSize limit" -ForegroundColor Red
			Write-Host "$(Get-Date): Please send $emailAttachment to $To" -ForegroundColor Red
		}
	}    

	Write-Verbose "$(Get-Date): Script started: $($Script:GPOStartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:GPOStartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
	}

	If($ScriptInfo)
	{
		$SIFile = "$($Script:pwdPath)\GetGpoBackupAndReportsScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName       : $ComputerName" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $Script:DevErrorFile" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Domain name        : $ADDomain" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "GPO Filter         : $GPOFilter" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $Log" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Organizational unit: $OrganizationalUnit" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $Script:RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
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
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

GetGpoBackupAndReports
#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up script"
#end of document processing

ProcessScriptEnd
[gc]::collect()
#endregion
