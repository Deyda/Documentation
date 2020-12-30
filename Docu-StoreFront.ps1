<#
.SYNOPSIS
	Creates documentation of a 3.x / 19xx StoreFront server cluster
	
.DESCRIPTION
	This PowerShell script uses the StoreFront PowerShell SDK to create StoreFront documentation.
	The script uses the PScribo documentation framework written by Iain Brighton
	(https://github.com/iainbrighton/PScribo) to render the output in any of the following formats:
		- MS Word
		- HTML
		- Text
	The script may used in command-line or GUI (default) mode. Command-line mode is useful for unattended
	documentation generation, while GUI mode expects input from the user before continuing. 
	The script must be run in an elevated PowerShell session (Admin mode).
	To run this script without regard to the current execution policy, execute the script as follows:
	> powershell.exe -executionPolicy bypass -file <directory>\Docu-StoreFront.ps1 <parameters>

.PARAMETER paramGUI
	Alias: GUI
	Use a graphical form to accept parameters from the user. 
	Note: GUI mode will also accept parameters passed on the command line.
	Default: True (use -GUI:$False to turn off)

.PARAMETER paramWord
	Alias: Word
	Generate the document in MS Word format.
	Default: True (use -Word:$False to turn off)

.PARAMETER paramHTML
	Alias: HTML
	Generate the document in HTML format.
	Default: True (use -HTML:$False to turn off)

.PARAMETER paramText
	Alias: Text 
	Generate the document in Text format.
	Default: False (use -Text to turn on)

.PARAMETER paramDirectory
	Alias: Dir
	Directory in which to place the generated documentation
	Default: Current working directory.
	
.PARAMETER paramFile
	Alias: FileName
	Base name of the generated output file(s). 
	Note: Extensions will be automatically be added for each document type generated.
	MS Word: .doc  HTML: .html  Text: .txt
	Default: <computer name> - StoreFront Documentation

.PARAMETER paramTitle
	Alias: Title
	Title of the document (placed on the first page)
	Default: StoreFront Documentation - 
		<computer name>

.PARAMETER paramAuthor
	Alias: Author
	Author of the document (placed on the first page)
	Default: $env:username
	
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	for each member in the server group.
	This parameter requires the script be run from an account with permission to retrieve hardware information.
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	Default: False (use -Hardware to turn on)

.PARAMETER Software
	Read the registry to obtain a list of installed software, as well as Citrix services installed 
	for each member in the server group.
	This parameter requires the script be run from an account with permission to read the registry.
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	You may exclude specific software by including their application names (wildcards allowed) in a file called
	SoftWareExclusions.txt in the same directory as the running script (see the sample file included). 
	Default: False (use -Software to turn on)
	
.EXAMPLE
	> powershell.exe -executionPolicy bypass -file c:\Scripts\Docu-StoreFront.ps1
	
	Will use all default values.
	
.EXAMPLE
	> powershell.exe -executionPolicy bypass -file c:\Scripts\Docu-StoreFront.ps1 -GUI:$False -FN "IPM-SF" -Dir "C:\Output" -Hardware
	
	Will gather hardware information for the host server, and
	will create the file(s) IPM-SF.<extension> in directory C:\Output without using the GUI.
	
.EXAMPLE
    > .\Docu-StoreFront.ps1 -Dir c:\temp -File StoreFront -Title "My SF Doc" -Author "Sam Jacobs"  -Text -Word:$False

    Will use the GUI and populate it with the base file name of "StoreFront", the title of "My SF Doc"
    the author "Sam Jacobs", and will add Text output to the default list, and remove MS Word output

.INPUTS
	None.  You cannot pipe objects to this script.

.OUTPUTS
	No objects are output from this script.  
	This script creates one or more files in the following formats: MS Word, HTML, and text.

.NOTES
	NAME:	  Docu-StoreFront.ps1
	VERSION:  4.0
	AUTHOR:  Sam Jacobs
	LASTEDIT: October 23, 2019
#>

Param(

	[parameter(Mandatory=$False)] 
	[Alias("Word")]
	[Switch]$paramWord=$True,

	[parameter(Mandatory=$False)] 
	[Alias("PDF")]
	[Switch]$paramPDF=$False,

	[parameter(Mandatory=$False)] 
	[Alias("Text")]
	[Switch]$paramText=$False,

	[parameter(Mandatory=$False)] 
	[Alias("HTML")]
	[Switch]$paramHTML=$True,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Hardware=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Software=$False,

	[parameter(Mandatory=$False)] 
	[Alias("Dir")]
	[string]$paramDirectory="",
    
	[parameter(Mandatory=$False)] 
	[Alias("FileName")]
	[string]$paramFileName="", 

	[parameter(Mandatory=$False)] 
	[Alias("Title")]
	[string]$paramTitle="", 

	[parameter(Mandatory=$False)] 
	[Alias("Author")]
	[string]$paramAuthor="", 

	[parameter(Mandatory=$False)] 
	[Switch]$paramDebug=$False,

	[parameter(Mandatory=$False)] 
	[Alias("GUI")]
	[Switch]$paramGUI=$True

	)

$SFDocVersion = "v4.0"

Write-Host ""
Write-Host "$(Get-Date): StoreFront Documentation Script $($SFDocVersion)"

# make sure script is running elevated
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "This script needs to be run elevated.`nPlease re-run this script as an Administrator!"
    Break
}

# process command line parameters & set defaults

Function setParamDefault($paramName, $defaultValue) {
	if ($paramName -eq "") { return $defaultValue } 
	else { return $paramName }
}

Function iif($paramToCheck, $valueIfTrue, $valueIfFalse) {
    if ($paramToCheck -eq $True) { return $valueIfTrue }
    else { return $valueIfFalse }
}

$paramDirectory =  setParamDefault $paramDirectory (Split-Path $script:MyInvocation.MyCommand.Path)
$paramFileName = setParamDefault $paramFileName "$($env:computername) StoreFront Documentation"
$paramTitle = setParamDefault $paramTitle "StoreFront Documentation - `r`n$($env:computername)"
$paramAuthor = setParamDefault $paramAuthor $env:username

$Script:OutputDir = $paramDirectory
$Script:OutputFile = $paramFileName
$Script:Title = $paramTitle
$Script:Author = $paramAuthor
$Script:OutputWord = $paramWord
$Script:OutputHTML = $paramHTML
$Script:OutputText = $paramText
$Script:Software = $Software
$Script:Hardware = $Hardware

Write-Verbose "$(Get-Date): Parameter OutputDir: $($Script:OutputDir)"
Write-Verbose "$(Get-Date): Parameter OutputFile: $($Script:OutputFile)"
Write-Verbose "$(Get-Date): Parameter Title: $($Script:Title)"
Write-Verbose "$(Get-Date): Parameter Author: $($Script:Author)"
Write-Verbose "$(Get-Date): Parameter OutputWord: $($Script:OutputWord)"
Write-Verbose "$(Get-Date): Parameter OutputHTML: $($Script:OutputHTML)"
Write-Verbose "$(Get-Date): Parameter OutputText: $($Script:OutputText)"
Write-Verbose "$(Get-Date): Parameter Software: $($Script:Software)"
Write-Verbose "$(Get-Date): Parameter Hardware: $($Script:Hardware)"
Write-Verbose "$(Get-Date): Parameter GUI: $($paramGUI)"
Write-Verbose "$(Get-Date): Parameter Debug: $($paramDebug)"

#region Documentation GUI
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#   GUI component for StoreFront server documentation script
#   Author:	 Sam Jacobs, IPM
#   Created:	 July, 2014
#   Version:	 3.0
#   Last Update: June 12, 2018 
#
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$continueProcessing = $True

#~~< GUI Customizations go here >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# title used for the form and any pop-up message boxes
$GUI_title  = "StoreFront Documentation Script $($SFDocVersion)"

#~~< Message Box buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[int]$MB_OK			= 0
[int]$MB_OK_CANCEL		= 1
[int]$MB_ABORT_RETRY_IGNORE 	= 2
[int]$MB_YES_NO_CANCEL		= 3
[int]$MB_YES_NO			= 4
[int]$MB_RETRY_CANCEL		= 5

#~~< Message Box icons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[int]$MB_ICON_CRITICAL		= 16
[int]$MB_ICON_QUESTION		= 32
[int]$MB_ICON_WARNING		= 48
[int]$MB_ICON_INFORMATIONAL	= 64

#~~< GUI Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function getDirectory($prompt) {
	$objShell = New-Object -com Shell.Application
	$selectedFolder = $objShell.BrowseForFolder(0,$prompt,0,0)
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objShell) | Out-Null
	return $selectedFolder
}

Function populateDirectory ($startDir) {
	$selectedDir = getDirectory("Please select the documentation output directory:")
	if ($selectedDir -ne $Null) {
		$txtOutputDir.Text = $selectedDir.Self.Path
	}
}

Function Abort_Script() {
	$Script:continueProcessing = $False
	$frmServer.Close()
}

Function Continue_Script() {
	# save fields needed from form before closing

	$Script:OutputFile = $txtOutputFile.Text
	$Script:OutputDir  = $txtOutputDir.Text
	$Script:Title = $txtTitle.Text
	$Script:Author = $txtAuthor.Text
	$Script:OutputWord = ($chkWord.Checked -eq $True)
	$Script:OutputHTML = ($chkHTML.Checked -eq $True)
	$Script:OutputText = ($chkText.Checked -eq $True)
	$Script:Software = ($chkSoftware.Checked -eq $True)
	$Script:Hardware = ($chkHardware.Checked -eq $True)
	
	if ( ($Script:OutputFile -eq "") -or ($Script:OutputDirectory -eq "") ) {
		[System.Windows.Forms.MessageBox]::Show("Output directory and filename cannot be null!" , 
			$GUI_title, $MB_OK, $MB_ICON_CRITICAL)
		Return
	}

	# make sure the output directory actually exists!
	If (!(Test-Path ($Script:OutputDir))) {
		[System.Windows.Forms.MessageBox]::Show("Output directory does not exist!" , 
			$GUI_title, $MB_OK, $MB_ICON_CRITICAL)
		Return
	}
	if ( (! $Script:OutputWord) -and (! $Script:OutputHTML) -and (! $Script:OutputText) ) {
		[System.Windows.Forms.MessageBox]::Show("Please select at least one output format!" , 
			$GUI_title, $MB_OK, $MB_ICON_CRITICAL)
		Return
	}
	$Script:continueProcessing = $True
	$frmServer.Close()
}

if ($paramGUI -eq $True) {
	Write-Verbose "$(Get-Date): Displaying GUI"
	#~~< create the GUI >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") |  Out-Null
	[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") |  Out-Null

	$frmServer = New-Object System.Windows.Forms.Form

	$btnExit = New-Object System.Windows.Forms.Button
	$btnGenerate = New-Object System.Windows.Forms.Button
	$groupBox3 = New-Object System.Windows.Forms.GroupBox
	$chkHardware = New-Object System.Windows.Forms.CheckBox
	$chkSoftware = New-Object System.Windows.Forms.CheckBox
	$groupBox1 = New-Object System.Windows.Forms.GroupBox
	$lblExt = New-Object System.Windows.Forms.Label
	$btnSelectOutputDir = New-Object System.Windows.Forms.Button
	$txtOutputDir = New-Object System.Windows.Forms.TextBox
	$txtOutputFile = New-Object System.Windows.Forms.TextBox
	$label4 = New-Object System.Windows.Forms.Label
	$label3 = New-Object System.Windows.Forms.Label
	$chkHTML = New-Object System.Windows.Forms.CheckBox
	$chkWord = New-Object System.Windows.Forms.CheckBox
	$chkPDF = New-Object System.Windows.Forms.CheckBox
	$chkText = New-Object System.Windows.Forms.CheckBox

	$label5 = New-Object System.Windows.Forms.Label
	$label6 = New-Object System.Windows.Forms.Label
	$txtTitle = New-Object System.Windows.Forms.TextBox
	$txtAuthor = New-Object System.Windows.Forms.TextBox

	## 
	## btnExit
	## 
	$btnExit.Location = New-Object System.Drawing.Point(268, 424)
	$btnExit.Name = "btnExit"
	$btnExit.Size = New-Object System.Drawing.Size(147, 30)
	$btnExit.TabIndex = 20
	$btnExit.Text = "Exit"
	$btnExit.UseVisualStyleBackColor = $True
	$btnExit.add_Click({Abort_Script})
	## 
	## btnGenerate
	## 
	$btnGenerate.Location = New-Object System.Drawing.Point(57, 424)
	$btnGenerate.Name = "btnGenerate"
	$btnGenerate.Size = New-Object System.Drawing.Size(147, 30)
	$btnGenerate.TabIndex = 19
	$btnGenerate.Text = "Generate"
	$btnGenerate.UseVisualStyleBackColor = $True
	$btnGenerate.Add_Click({Continue_Script})
	## 
	## groupBox3
	## 
	$groupBox3.Controls.Add($chkHardware)
	$groupBox3.Controls.Add($chkSoftware)
	$groupBox3.Location = New-Object System.Drawing.Point(36, 319)
	$groupBox3.Name = "groupBox3"
	$groupBox3.Size = New-Object System.Drawing.Size(403, 88)
	$groupBox3.TabIndex = 18
	$groupBox3.TabStop = $False
	$groupBox3.Text = " Optional "
	## 
	## chkHardware
	## 
	$chkHardware.AutoSize = $True
	$chkHardware.Location = New-Object System.Drawing.Point(29, 26)
	$chkHardware.Name = "chkHardware"
	$chkHardware.Size = New-Object System.Drawing.Size(217, 17)
	$chkHardware.TabIndex = 4
	$chkHardware.Text = "Use WMI to gather hardware information"
	$chkHardware.UseVisualStyleBackColor = $True
	$chkHardware.Checked = ($Hardware -eq $True)
	## 
	## chkSoftware
	## 
	$chkSoftware.AutoSize = $True
	$chkSoftware.Location = New-Object System.Drawing.Point(29, 55)
	$chkSoftware.Name = "chkSoftware"
	$chkSoftware.Size = New-Object System.Drawing.Size(217, 17)
	$chkSoftware.TabIndex = 3
	$chkSoftware.Text = "Query registry for installed software"
	$chkSoftware.UseVisualStyleBackColor = $True
	$chkSoftware.Checked = ($Software -eq $True)
	## 
	## groupBox1
	## 
	$groupBox1.Controls.Add($lblExt)
	$groupBox1.Controls.Add($btnSelectOutputDir)
	$groupBox1.Controls.Add($txtOutputDir)
	$groupBox1.Controls.Add($txtOutputFile)
	$groupBox1.Controls.Add($label4)
	$groupBox1.Controls.Add($label3)
	$groupBox1.Controls.Add($label5)
	$groupBox1.Controls.Add($label6)
	$groupBox1.Controls.Add($txtTitle)
	$groupBox1.Controls.Add($txtAuthor)
	$groupBox1.Controls.Add($chkWord)
	$groupBox1.Controls.Add($chkHTML)
	$groupBox1.Controls.Add($chkText)
	$groupBox1.Location = New-Object System.Drawing.Point(34, 92)
	$groupBox1.Name = "groupBox1"
	$groupBox1.Size = New-Object System.Drawing.Size(405, 216)
	$groupBox1.TabIndex = 21
	$groupBox1.TabStop = $False
	$groupBox1.Text = " Output "

	## 
	## btnSelectOutputDir
	## 
	$btnSelectOutputDir.Location = New-Object System.Drawing.Point(338, 28)
	$btnSelectOutputDir.Name = "btnSelectOutputDir"
	$btnSelectOutputDir.Size = New-Object System.Drawing.Size(37, 20)
	$btnSelectOutputDir.TabIndex = 10
	$btnSelectOutputDir.Text = "..."
	$btnSelectOutputDir.UseVisualStyleBackColor = $True
	$btnSelectOutputDir.Add_Click({populateDirectory($OutputDir)})
	## 
	## txtOutputDir
	## 
	$txtOutputDir.Location = New-Object System.Drawing.Point(114, 28)
	$txtOutputDir.Name = "txtOutputDir"
	$txtOutputDir.Size = New-Object System.Drawing.Size(215, 20)
	$txtOutputDir.TabIndex = 9
	$txtOutputDir.Text = $paramDirectory
	## 
	## txtOutputFile
	## 
	$txtOutputFile.Location = New-Object System.Drawing.Point(114, 60)
	$txtOutputFile.Name = "txtOutputFile"
	$txtOutputFile.Size = New-Object System.Drawing.Size(215, 20)
	$txtOutputFile.TabIndex = 8
	$txtOutputFile.Text = $paramFileName
	## 
	## label4
	## 
	$label4.AutoSize = $True
	$label4.Location = New-Object System.Drawing.Point(25, 28)
	$label4.Name = "label4"
	$label4.Size = New-Object System.Drawing.Size(49, 13)
	$label4.TabIndex = 1
	$label4.Text = "Directory:"
	## 
	## label3
	## 
	$label3.AutoSize = $True
	$label3.Location = New-Object System.Drawing.Point(25, 61)
	$label3.Name = "label3"
	$label3.Size = New-Object System.Drawing.Size(49, 13)
	$label3.TabIndex = 0
	$label3.Text = "Filename:"
	## 
	## lblExt - formats
	## 
	$lblExt.AutoSize = $True
	$lblExt.Location = New-Object System.Drawing.Point(25, 94)
	$lblExt.Name = "lblExt"
	$lblExt.Size = New-Object System.Drawing.Size(34, 15)
	$lblExt.TabIndex = 11
	$lblExt.Text = "Formats:"
	## 
	## chkWord
	## 
	$chkWord.AutoSize = $True
	$chkWord.Location = New-Object System.Drawing.Point(114, 94)
	$chkWord.Name = "chkWord"
	$chkWord.Size = New-Object System.Drawing.Size(75, 17)
	$chkWord.TabIndex = 4
	$chkWord.Text = "MS Word"
	$chkWord.UseVisualStyleBackColor = $True
	$chkWord.Checked = ($paramWord -eq $True)
	## 
	## chkHTML
	## 
	$chkHTML.AutoSize = $True
	$chkHTML.Location = New-Object System.Drawing.Point(209, 94)
	$chkHTML.Name = "chkHTML"
	$chkHTML.Size = New-Object System.Drawing.Size(75, 17)
	$chkHTML.TabIndex = 4
	$chkHTML.Text = "HTML"
	$chkHTML.UseVisualStyleBackColor = $True
	$chkHTML.Checked = ($paramHTML -eq $True)
	## 
	## chkText
	## 
	$chkText.AutoSize = $True
	$chkText.Location = New-Object System.Drawing.Point(284, 94)
	$chkText.Name = "chkText"
	$chkText.Size = New-Object System.Drawing.Size(75, 17)
	$chkText.TabIndex = 4
	$chkText.Text = "Text"
	$chkText.UseVisualStyleBackColor = $True
	$chkText.Checked = ($paramText -eq $True)
	## 
	## label5
	## 
	$label5.AutoSize = $True
	$label5.Location = New-Object System.Drawing.Point(25, 144)
	$label5.Name = "label5"
	$label5.Size = New-Object System.Drawing.Size(49, 13)
	$label5.TabIndex = 11
	$label5.Text = "Title:"
	## 
	## txtTitle
	## 
	$txtTitle.Location = New-Object System.Drawing.Point(114, 144)
	$txtTitle.Name = "txtTitle"
	$txtTitle.Size = New-Object System.Drawing.Size(215, 20)
	$txtTitle.TabIndex = 9
	$txtTitle.Text = $paramTitle
	## 
	## label6
	## 
	$label6.AutoSize = $True
	$label6.Location = New-Object System.Drawing.Point(25, 177)
	$label6.Name = "label6"
	$label6.Size = New-Object System.Drawing.Size(49, 13)
	$label6.TabIndex = 11
	$label6.Text = "Author:"
	## 
	## txtAuthor
	## 
	$txtAuthor.Location = New-Object System.Drawing.Point(114, 177)
	$txtAuthor.Name = "txtAuthor"
	$txtAuthor.Size = New-Object System.Drawing.Size(215, 20)
	$txtAuthor.TabIndex = 9
	$txtAuthor.Text = $paramAuthor

	## 
	## frmServer
	## 
	$frmServer.ClientSize = New-Object System.Drawing.Size(475, 483)
	$frmServer.Controls.Add($groupBox1)
	$frmServer.Controls.Add($btnExit)
	$frmServer.Controls.Add($btnGenerate)
	$frmServer.Controls.Add($groupBox3)
	$frmServer.Name = "frmServer"
	$frmServer.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
	$frmServer.Text = $GUI_title

	#region formIcon

		# form icon & logo - convert to base64
	
[string] $iconBase64=@"
iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAMAAABHPGVmAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAkZQTFRF//////7+/PXu+ejY9tzD9Na489Kw8s2p8cyl8s+r89Oz9di79+DK+u3g/fn2/vr2++7i9dm97bmF5p5X4pE+34Yr3XsZ23YP2nIJ2W4D2W0A2W8F2nML23cS3X8f4Ioy45VG6KZl8MWa9+PO/PPp/v37//79+u7i8cul56Vi34Qp23ML2GsA2GoA2WsA2WwA2nAF4o8767N6/ffx9+DJ6q904Igv23MM2WwB3HkW45ND78KV+u3f//389t3E6alp3Xwb2m8E4Ioz7ryL///++ena6rB23X0c2nAG4Ys08Mig/fbw/fjz8cmh4Io02W0B99/I+u7h6att23QN3oIl8Med9t3F45NC2W4E23UP6a1x+/Dl5p5W+uze/vz63oAh45ZI+OXS3HoY4pJB5JlM+uzf56Jd/PPr2nAH7LaA/fn189Ct4Iky+enY/vz5/v389Ne65ZpP/PTt/vv323QM89Gw5ZtR3HgV9Ne56q5y/vz74IYt+u3h2nEI8s6p6apq4pA+/PTs2W4C3X0d99/H2W8D8s2n4pE/7buJ6Kho/vv45JdK/ffy9Na34Y054Icu3oEj/PPq3X4f/PLp3oMn/PXt4pA9/fbv5Z1U/fn089Oy+efV5JhM7LV/23YQ9tvB9di83HkX45RF7r2M78KU6q9z2m8F3X4e9+LM4Ys12W0C/vv58Mif+enZ/Pbv2nEJ8MWb9t/I3oIm5qFb+ObU3HoX78CS2nIK7LeC/vr32nIL7LeB89Kx2W0D45RE//782W8E34QoubFMPAAAAAFiS0dEAIgFHUgAAAONSURBVGje7dn5O1RRGAfwM0OyZEtkrhmkbGPMsTSTscvWIiKKwUS21FRIlqyDJEnJUlKppKREifbtP+uOenomy7nnnnP7oaf7/QPu53nfszznnAuAGDFixIj5NyKRWllvstlsa2fvsMXR6S8Azi6uW922uXts95QxXnKFt4/vDr+du/wFFAICg4KVITJVqBpChg2EUBUKwxThEa6RQhWxW6Pdo175+qrAKF10TGwcPeEfn5DIrEusMOqkvcmOKZRGalp61EbCL2ff/gMHaYiMQ5lZaMIctdztcACxkZ1zhJtYqSY3z5mMCDh6LB/PYJUCfSGJkWKvVGMSZkVWVMzfMAQdxy3jZ2QlpbzrCDrBz2CLKbPiZ0jsedZhTlZ5BS8kXsnfYJXKKh5GdjSPMbfIyWr8LfNUjozIYKDuNDZyCHMNro3aeAbTSM0kNdhh0Z/FMs6lERPmhtVgIdbp5IWwSm0dhnE+gcZgmPoLGMjuRDoENmRwGhINFcEijU2cSKCWbB1aKM0XuZCgELpusUhLK4fhHExbCMO0tXMgLkQ746pSOjgQ1xBqg1F1mtDI1i56BHb3oCewm4oeYS71IhHp5VABENiHRKzc6ScXi1zpRyHWHvSTi0WuDqCQawWCIINSFGLjKYDBwOvIObxZgBnMIjeGUIitEAYDbw6jEDsvQZAR5PnLXi7IwBtHUYiDQhBk7BYK2eItCHJ7HIU4+gixd0EN8rTq5CvI3hUDkNkhQLug/A4a8QujR9TKCTSyU0GPqO5yXIWHw+n71XUPbYC4CHrk/iQHAh48pD53PTJxIZGENzmLbk1xGQA8TqIsJBfjPh9Lcc9ayTTOY07yE6pCns5gGCD2Gc2ozGqe4yDghZyiEO0clgFeuhGPCpxfwDMAePWaWFkcxUUkeYTHL2h8g2sAcEpPdDSCSzg3398pLCJQYMGygQ8Cikt4Nwy+ffeelwFAadksT+ND2gBPg71FlPNSYMFH/gYAFZVh+C2DS8tkL8NV1TrMDQbOGj/hPUGtjeG0cRanGDi/OEFIsJGc0es4FfhZu4C9ztdNf01tPUQ5MKtbMyehMtjU2TQ0bsjA/NzpmX5awpyMpuaWtijVakit6rr/Zar0qxCEOYbW9o7O7m/QMvLvd+9Nmqgb9UdSTD29fVeuDl6/cXPEOHZbE3NngvCXCUck/QNS09Bw1eitcSF//okRI0bM/5sfl10eGtmsheUAAAAldEVYdGRhdGU6Y3JlYXRlADIwMTQtMTEtMTBUMjI6NDI6MjQrMDA6MDCQj02lAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDE0LTExLTEwVDIyOjQyOjI0KzAwOjAw4dL1GQAAAABJRU5ErkJggg==
"@
		$iconStream=[System.IO.MemoryStream][System.Convert]::FromBase64String($iconBase64)
		$iconBmp=[System.Drawing.Bitmap][System.Drawing.Image]::FromStream($iconStream)
		$iconHandle=$iconBmp.GetHicon()
		$icon=[System.Drawing.Icon]::FromHandle($iconHandle)
		$frmServer.icon = $icon

[string] $logoBase64=@"
iVBORw0KGgoAAAANSUhEUgAAAFAAAAA1CAYAAADWKGxEAAAACXBIWXMAABYlAAAWJQFJUiTwAAAAB3RJTUUH4gURAwQk7TQwLwAAAAd0RVh0QXV0aG9yAKmuzEgAAAAMdEVYdERlc2NyaXB0aW9uABMJISMAAAAKdEVYdENvcHlyaWdodACsD8w6AAAADnRFWHRDcmVhdGlvbiB0aW1lADX3DwkAAAAJdEVYdFNvZnR3YXJlAF1w/zoAAAALdEVYdERpc2NsYWltZXIAt8C0jwAAAAh0RVh0V2FybmluZwDAG+aHAAAAB3RFWHRTb3VyY2UA9f+D6wAAAAh0RVh0Q29tbWVudAD2zJa/AAAABnRFWHRUaXRsZQCo7tInAAAQiElEQVR4nO1aeXRUVZr/3bfXkqpKUtkDBKIsItBAywDNINN0uyAH6AahndFWsWWRNDoDHWhBFlmbBoGWtqcRW0ZHUHDEBVpslU0ZFwQBJSARCSFAyFKVqtSrevudP5JKqkioJAQHOYffOe/Uqfu+736/+93tu/d7hFJK1ZJD8P99NYInP4MzszMcmTetcT24bg4hJIwbSAiilBymZavHQo2EwdhcoIaGzuluVEvZW3MKd4y/1gR/6GB8O1Y1OA8ACCeg1CfD9J29t2L9w4OuMb8fPJjgN580OC8KgxIk2wVAcN4YgS2AcaZ3ADX0uEJCKUAIOF4ou0a8rhswfEqH9dnJNlDLAqUUlFJ4JAaVYeu0ItON15rgDx1c+s0jf1etBn6UxZ4dQC0TBBSW5LrAZ936iPfhZ6qvNcEfOgilFFWUOlNfemSIZtK7NcqUy0LyhswHVlZea3LXAwil9FpzuK7BXGsC1ztuOLCduOHAduKGA9sJbvfu3eyBAwfSASAcrrs7sNvtkGXZt3DhQvVyioWFhVler7dBx+v1QhCEmkmTJkWiMhs2bOh15uzZn/urq7MyMzPv0A0D5RcubBNFMdCpU6ftM2bMKG4r4U2bNnl8Pp8tXFWF6E2H1+tFenq6f/z48QoAzJs37y6/39/L4/GMI4QIPp9vq91uPzZixIg9w4YNCzRX74YNG/K//vrrkaIo9vUkJ/eVQ6Hg+fLybR6XK5ifn//OtGnTLjanR5YsWdKtoqLioCzLDYU2mw2pqamL58+fv7w5pXXr1j166tSp1aFQCABAKYXb7YbT6SxYuHDhxkWLFnWtrK7+DyUSeYhjWVFRFBiGAUIIOI4Dx3GwLEsVRXGf1+v9w4IFCz5srQNXr1nzwpkzZyaEQiGQetvJycmkY8eOwzI6dar8ZO/eF0PB2mEAhaZpAABBEAAAkiSVuVyux5csWfJGtL4VK1ZkVVRVzYvI8sO6YYi6psE0TTAMA47jQRgCnuNqk5KSlo4ZM2bNoEGDlFg+7ODBgzNqa2ufkGVZME1TMAxDACDU1NS8OHr06K+aa8SRI0ceKC8vHxoMBgXTNAVd1wWWZQWPx/POrl27upaVlb1lGMaQSDjMKYoCK+aUY5omNF2HYZqcZVn5siz/et/HHzv79e37id1u11py4Ef79v2qorKyXyQcEUzTEHRNExiW5R0Ox3eHv/hilRKJ9A+HwzAMo8GmYRgwDAOqqroIYSYcOnRQGjp06IeTJ0/uEAgEdkfC4TtDoRCnaVoDV8uy6vR0HaqqipTiZ999d2qg2+1+Mzc3t4Enw7IsjRqKfSzLMi7XCEqpYtXHj9E4MhwOIxAIzC4vL381oiiecP2IJoQ00ScAQAFN0yDLMsKh0MzVq1e/VVxcnNSSAwkhZgNPUBCGQUiWcfr06WWyLHcNh8PN2oxyrfZVo6zs3OzCwsKpmqbtUFX15nA40gxXEvMAwWAA4XD4Z6+99lphbJ1XvonEBOCEEOi6josXL94cCoXq3hECluMgimLcw7JctIIGXb/fD7+/5qfPb9iwttX2SUMVME0TtbW1ME0LDMNAEEWIogRRFJuoMYSgtjYIn9/3HKXopaoaQACWZSGKIqR6ngzDgFKrwQjDMAgGgwgEg1MWLJiZ1lBf27x2+cYQQmBRq+4vIbDbbGBZ9mh2dvZOb3r6E1lZWXMzMjJ2EoYUS5KE2BNQHbkAqqurH5w+ffrIVtmM7cD6X57nwfOCz5uWtjM3N2dGdnb2q263289xXJw9QggiEQWGoQOgEAUBhGVL09LSdrrc7l97vd4/JyU5T9tstjg9i1rgWNaracJj0TIOTdD88E/cmMY2sQwDp8tVm5mV9dTYMWOe69mzZ9xd2YoVK5wVFRWzbTbbHJ/PFzdtDF1ndN1cs3nz5s/vu+++irZQYFgWaelp5V06d7576tSph6PlGzdu7LR///4thJABut5IJWqV4zg4HI5vMjp1umPuzJml9cUvb9++PXvXrl3/MAyjp2EY9ToEumHA5/P1qa+CNjMCW3s2buponuPAC0KZXZL+ad6cOWsvdR4AFBYWhlauXDk3Nzd3idvtjuthwzDAMMg/ceLEpJbNx9uXJMlwp6SMj3UeADz00ENnOuTlzRcEAZce+ymlSEpKQrdu3SbHOA8AMHLkyPMej+c3giDE5YU0VYUoisP27NmTAsRN4baOvKaOZllWy8rJ+f3KlSuPt6Q9a9asuQzD/FmSbI0MCIEsy6iqqhpTUlLSdAG7DARBAMey++fOnv1Rc+/nzZmzU9f1L3k+fsLxggDDNN8rKCjY26zevHmfsizrY5hGNzGEQSAQYN9+++26/43i7buV4TgONofjy8ULF/53a3W6d+8+j+e5QOxoMk0TkUik5+uvv+5KoBq3BjIMg/TMzEOJxNPS02sJiW+uJElIcjqbDdWiVbvc7tK43ZkQmKZJy8vL6wQaXySk2xQxlVJKIdlsSPF43mpLFQUFBT5JkvbxbOPIsCwLgiDwkUjkrtbWQwGUlJa+m0gmLMvbWDZ+wpmGgZSUlEQOtMrOnXuf5/lLrBEkJdVFXExceVtwyQjQNE12OBzb2lgLXE7nAcnWuCsTQmBZFltSUpKcULG+AymlEHgeGdnZQiJxv9/vi1cnUFUVhw4d+jShGUr5pqUxbU9IspUghIGuadr+/fvbnoRimHcjkQhi1xnLsqBp2mUDeQANHUgpBcuy8CYljsEJxzUTcQCSJEkt8Ev8OuHbVoOC4zjSs2fPhKOgORw9erTSNM0m5SzLXhVm7QUDNNnxm7yvwxXEfzHQDYOWlpY2CVtaQu/evdManBVDoTmnxiEqS0jCBl4VXBr/xNi7KruwaZpISkpyDh48eEhbdS3Luttmt8O6ZE0VBKHZKdcAWt8ISps2sLVoheOtZm03uwa2rxct0+SLi4t7tVXP7/ffpkQiDRcMlFIwDGPm5eX5E2tehWRYK6poaY278jAmRoEQgoiiQNO0iceAZnat5rFq1apcVdOGRo9KQP2Oruu6zWbb2bJ9EiXQJuaNaGsnNLUTE8a0OY6J+2eaJnTD6PbCzJkzWltDSUnJBmpZ7viLBRYOu/2rcePGBVu2T7//9a+JzXhctZwIARCWZWiKsuyPq1YVtCQ/ffr0xeFw+E5VjWYNCCilcDjsYBhmVV5e3mXTCXXiV+K4q+/sxAt1G0EIQU1NDYzi4mdnzZqV37lz5zVTpkw5Eyuy9rnn+hcdOVJgmuaD0Sv3OlCwHAeW406NGjVqd6uNtmkTuVSONFPWNjQ6sM092lSeUgpCCGpDIViW9UR1dfWDEydO/CwlNRWWaUJRVC9Af6wbBjRVbaJrt9ngsNufGD58eMtXWTRm+v6/TuN4NDowllCreqWpjCRJ0HUdpmlBlmUwDJPM8/xdfl/dKSqamyCxDa8fPampqbAsa87y5cu3t4p5rNOueBBd+dVdFJdZA9vWo5TSaCbvvyRJKouejiil0DQdqqpCVVWYphmfd6g/hnEcB5fH8+Kzzz67sk2GgfpAus1abQRttHUJGJZlWVGSmuQuGIaxNZFOAEEQINhs2/v06XOnw2F/3+V21eUVQOMuTSkap7rNZoPd6fQlJyc/vvjppycSQlrMyhFCnJIkQRSEOq6CgPoUQcJjJCFEvLSNoiiCUprwzEgptTfICw2+cYuiSACAUxTFp6jqJsMwGg7OrK6jS5cuzQbhCQyBGkbqtGnTigDc8+STT06WRGlaJKJ0Z1kG0d02ejVECKlgOO4Fj8v1n4sXLy5NUHUcDMN4X9f1cB1fAFbdLTGl9HQiPYFliwzD2KTrOsAwYOo4gFKa8BtIQsgnhmGk16UDGLAsASEkEgqFFOAKP29bv379sqNHj872+/1REnC73fB4PJOXLl26Pip3/vx5244dOwaWl5f3PVFcDJgmOnTogJyOHctEjvtw0qRJ1/0HnO0IY1peeLKzsyOPPvrobgCtD0uuM7QjkL7xYSZw4+usduOqnkR+KFiyZEmfoqIi/63duqX37t+fOX78eLlhGK5AICB3797dSwzDOvjFUWXtX9YemzFjRrectBy+rKLs3DPPPNPCDVBTtN+B7T8NXSlY1F3XNVinlNoLCgqIxPMv3da/f4bH40net2ePGg5HvvF43EWqonS6cO78j3mOO9ZvQN8BixYuXEQs+mhF9YVT6QhNAXBZB1JK7YQQFUDcTW/7HXgNnEcG/rvUu1/3zwSeO3dg7SMjAGD4rJeXRoh47/Bbb7+n8vN3x+pa5NZvv/NV73j73XBu51xryMCBzoy8vPyjR44sSkpKIpppZOuKUsQx3Oke2UmffYShg/5t+f8kvzJ7bJPcsmfU0l433Zz3LtUj9x/80yN7Yt9dl1P4vl8OHbDlo2968zzXe9y8zQNef/q+zw+ePPtz1USH/13+5ElMHQ8A3wLA4sWLY1WbS7x/0ffxF+eVnDsx518G5ufWl3EARAAyAGR4nRO+Lgvm9MsQGw7wU9du75me6t56XTqwuPTibZwgQtEMfFxUOhLACYfT1ZutCWx8fMM/+m3Z9+0ct8iMOlNZs2L8T255desXpQ9Ftry3dN+hZ8i4NXvnp0jmt8f/Onldx/vXLquRtTyOo/fWVAaVbe8duXPI+ZqQT6F/Mg3D6wuGnqp8dcYqWdFvN0PB4tuH9T0ujF2zVo/I8v0/7bW+I8+9dF068HxVqKvXaYNAuHNVgdppI556cUuFL8B2y/GYW3cf26HpVnmqN3nd+YD45AeHTnbUayP3/uS3v/h00aufZlZrTIGHVfY+8Mc3TgRg+x1o+DP/xUDktr5dbYNvyf3bX7cftLwe+5upSWJQtviVoxe8ogVCCpfiFKTnPizaxoMbLEn8uJdnjioBsPy6DGNUwxxhhENf9u+SviRiMCnFZbV/sAmMmZ+V8pvKgJKZn+E2RJ4XXTYRF/y1Hbxuew0vMWP3Hy/9hWiGoRu6/Ob+ExNZJXjy978asogRHeZjI/uDZ1le8QdIx3SXkmQTBQYUnx47+88cR26q8CkddM45rHuG+ETwjcJ3olyuyIGkDrj0oZR+7/ciKzbvz5It1pnqcVZunTdhY4ZbUM9cqB2RmewQHDaBN2oDxvGyiqIvT12QgiF5a++8jN0pLnvNkdPVdzEs09nFKs9Xy9oQyknjOyYLczfu/DKZdbucffMzcPi7C+BELlJUUmkWX/DbJcbcmul11wYViLnpjgo1FCgvqwzdGsvnSkcgy7IsGIYBwzJgWBZMXW73ex/Rr+072stkJY8ksh8DiKS5HK/pYRU5qS6MHtQDgACqG+/UXKzZaLc7hvS/pdPeouKyN/1hxt01x7tPVdW/ByOMi7O0M4efn/66XeJ/aYYjKD7nQ48OaTAMWD1y3a+k2QV/itvZu2dexnhLN6Up9/xoZOdU6WhVxJgwasVbziifK1oDDcNYoQEvAAC16uIYsy6zVt5uD7WAqpChD+rRCb3Sma8AwCnyC/Jvyh55e58uKRNuvwVnfns3t2zT/q2p2UnI8kjLX5h+z96vvikdzbm8GNkrbenBU+5+5aoIM1Q1HwBN8Tjfks7V3LNo80fCB8vuR0Qzkv+28/CHXXNS8ZfH7sIHR87gth6Z4Tn/OvzAwZMXt3izxTt85yv7A9gLAP8Hj5ZTZfsLQT8AAAAASUVORK5CYII=
"@
	$imageBytes = [Convert]::FromBase64String($logoBase64)
	$ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
	$ms.Write($imageBytes, 0, $imageBytes.Length);
	$logo = [System.Drawing.Image]::FromStream($ms, $true)
 
	$pictureBox = new-object Windows.Forms.PictureBox
	$pictureBox.Width =  101
	$pictureBox.Height =  66; 
	$pictureBox.Location = New-Object System.Drawing.Size(190,20) 
	$pictureBox.Image = $logo;
	$frmServer.Controls.Add($pictureBox)

	#endregion

	# display the form
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($frmServer)

	If ($continueProcessing -eq $False) { 
		Write-Verbose "$(Get-Date): Script cancelled by user."
		Return 
	}
	#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	#  End of GUI component 
	#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
}
#endregion Documentation GUI

Write-Verbose "$(Get-Date): Output directory: $($Script:OutputDir)"
Write-Verbose "$(Get-Date): Output file: $($Script:OutputFile)"
Write-Verbose "$(Get-Date): Title: $($Script:Title.Replace("`r`n"," "))"
Write-Verbose "$(Get-Date): Author: $($Script:Author)"
Write-Verbose "$(Get-Date): Doc outputs: $(iif $Script:OutputWord 'Word ' '')$(iif $Script:OutputHTML 'HTML ' '')$(iif $Script:OutputText 'Text ' '')"
Write-Verbose "$(Get-Date): Misc outputs: $(iif $Script:Software 'Software ' '')$(iif $Script:Hardware 'Hardware ' '')"

# ~~~~~~~~~~~~~  PScribo bundle  ~~~~~~~~~~~~~~~

#region PScribo Bundle v0.7.21.110
#requires -Version 3

<#
    The MIT License (MIT)

    Copyright (c) 2018 Iain Brighton

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>

$localized = DATA {
# en-US
ConvertFrom-StringData @'
ImportingFile                    = Importing file '{0}'.
InvalidDirectoryPathError        = Path '{0}' is not a valid directory path.'
NoScriptBlockProvidedError       = No PScribo section script block is provided (have you put the open curly brace on the next line?).
InvalidHtmlColorError            = Invalid Html color '{0}' specified.
InvalidHtmlBackgroundColorError  = Invalid Html background color '{0}' specified.
UndefinedTableHeaderStyleError   = Undefined table header style '{0}' specified.
UndefinedTableRowStyleError      = Undefined table row style '{0}' specified.
UndefinedAltTableRowStyleError   = Undefined table alternating row style '{0}' specified.
InvalidTableBorderColorError     = Invalid table border color '{0}' specified.
UndefinedStyleError              = Undefined style '{0}' specified.
OpenPackageError                 = Error opening package '{0}'. Ensure the file in not in use by another process.
MaxHeadingLevelWarning           = Html5 supports a maximum of 6 heading levels. Reduce the number of nested Document sections to remove the unsupported tags in the resulting Html output.
TableHeadersWithNoColumnsWarning = Table headers have been specified with no table columns/properties. Headers will be ignored.
TableHeadersCountMismatchWarning = The number of table headers specified does not match the number of specified columns/properties. Headers will be ignored.
ListTableColumnCountWarning      = Table columns widths in list format must be 2. Column widths will be ignored.
TableColumnWidthMismatchWarning  = The specified number of table columns and column widths do not match. Column widths will be ignored.
TableColumnWidthSumWarning       = The table column widths total '{0}'%. Total column width must equal 100%. Column widths will be ignored.
TableWidthOverflowWarning        = The table width overflows the page margin and has been adjusted to '{0}'%.
UnexpectedObjectWarning          = Unexpected object in section '{0}'.
UnexpectedObjectTypeWarning      = Unexpected object '{0}' in section '{1}'.

DocumentProcessingStarted        = Document '{0}' processing started.
DocumentInvokePlugin             = Invoking '{0}' plugin.
DocumentOptions                  = Setting global document options.
DocumentOptionSpaceSeparator     = Setting default space separator to '{0}'.
DocumentOptionUppercaseHeadings  = Enabling uppercase headings.
DocumentOptionUppercaseSections  = Enabling uppercase sections.
DocumentOptionSectionNumbering   = Enabling section/heading numbering.
DocumentOptionPageTopMargin      = Setting page top margin to '{0}'mm.
DocumentOptionPageRightMargin    = Setting page right margin to '{0}'mm.
DocumentOptionPageBottomMargin   = Setting page bottom margin to '{0}'mm.
DocumentOptionPageLeftMargin     = Setting page left margin to '{0}'mm.
DocumentOptionPageSize           = Setting page size to '{0}'.
DocumentOptionPageOrientation    = Setting page orientation to '{0}'.
DocumentOptionPageHeight         = Setting page height to '{0}'mm.
DocumentOptionPageWidth          = Setting page width to '{0}'mm.
DocumentOptionDefaultFont        = Setting default font(s) to '{0}'.
ProcessingBlankLine              = Processing blank line.
ProcessingImage                  = Processing image '{0}'.
ProcessingLineBreak              = Processing line break.
ProcessingPageBreak              = Processing page break.
ProcessingParagraph              = Processing paragraph '{0}'.
ProcessingSection                = Processing section '{0}'.
ProcessingSectionStarted         = Processing section '{0}' started.
ProcessingSectionCompleted       = Processing section '{0}' completed.
PluginProcessingSection          = Processing {0} '{1}'.
ProcessingStyle                  = Setting document style '{0}'.
ProcessingTable                  = Processing table '{0}'.
ProcessingTableStyle             = Setting table style '{0}'.
ProcessingTOC                    = Processing table of contents '{0}'.
ProcessingDocumentPart           = Processing document part '{0}'.
WritingDocumentPart              = Writing document part '{0}'.
GeneratingPackageRelationships   = Generating package relationships.
PluginUnsupportedSection         = Unsupported section '{0}'.
DocumentProcessingCompleted      = Document '{0}' processing completed.
                                 
TotalProcessingTime              = Total processing time '{0:N2}' seconds.
SavingFile                       = Saving file '{0}'.

IncorrectCharsInPath             = The incorrect char found in the Path.
IncorrectCharsInName             = The incorrect char found in the Name.
'@;
}

function BlankLine {
<#
    .SYNOPSIS
        Initializes a new PScribo blank line object.
#>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline, Position = 0)]
        [System.UInt32] $Count = 1
    )
    begin {
        #region BlankLine Private Functions
        function New-PScriboBlankLine {
        <#
            .SYNOPSIS
                Initializes a new PScribo blank line break.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(ValueFromPipeline)]
                [System.UInt32] $Count = 1
            )
            process {
                $typeName = 'PScribo.BlankLine';
                $pscriboDocument.Properties['BlankLines']++;
                $pscriboBlankLine = [PSCustomObject] @{
                    Id = [System.Guid]::NewGuid().ToString();
                    LineCount = $Count;
                    Type = $typeName;
                }
                return $pscriboBlankLine;
            }
        } #end function New-PScriboBlankLine
        #endregion BlankLine Private Functions
    } #end begin
    process {
        WriteLog -Message $localized.ProcessingBlankLine;
        return (New-PScriboBlankLine @PSBoundParameters);
    } #end process
} #end function BlankLine

function Document {
<#
    .SYNOPSIS
        Initializes a new PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## PScribo document name
        [Parameter(Mandatory, Position = 0)]
        [System.String] $Name,
        ## PScribo document DSL script block containing Section, Paragraph and/or Table etc. commands.
        [Parameter(Position = 1)]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),
        ## PScribo document Id
        [Parameter()]
        [System.String] $Id = $Name.Replace(' ','')
    )
    begin {
        $pluginName = 'Document';
        #region Document Private Functions
        function New-PScriboDocument {
        <#
            .SYNOPSIS
                Initializes a new PScript document object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseLiteralInitializerForHashtable','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## PScribo document name
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,
                ## PScribo document Id
                [Parameter()]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = $Name.Replace(' ','')
            )
            begin {
                if ($(Test-CharsInPath -Path $Name -SkipCheckCharsInFolderPart) -eq 3 ) {
                    throw -Message ($localized.IncorrectCharsInName);
                }
            }
            process {
                WriteLog -Message ($localized.DocumentProcessingStarted -f $Name);
                $typeName = 'PScribo.Document';
                $pscriboDocument = [PSCustomObject] @{
                    Id = $Id.ToUpper();
                    Type = $typeName;
                    Name = $Name;
                    Sections = New-Object -TypeName System.Collections.ArrayList;
                    Options = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    Properties = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    Styles = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    TableStyles = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    DefaultStyle = $null;
                    DefaultTableStyle = $null;
                    TOC = New-Object -TypeName System.Collections.ArrayList;
                }
                $defaultDocumentOptionParams = @{
                    MarginTopAndBottom = 72;
                    MarginLeftAndRight = 54;
                    PageSize = 'A4';
                    DefaultFont = 'Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif';
                }
                DocumentOption @defaultDocumentOptionParams -Verbose:$false;
                ## Set "default" styles
                Style -Name Normal -Default -Verbose:$false;
                Style -Name Title -Size 28 -Color 0072af -Verbose:$false;
                Style -Name TOC -Size 16 -Color 0072af -Hide -Verbose:$false;
                Style -Name 'Heading 1' -Size 16 -Color 0072af -Verbose:$false;
                Style -Name 'Heading 2' -Size 14 -Color 0072af -Verbose:$false;
                Style -Name 'Heading 3' -Size 12 -Color 0072af -Verbose:$false;
                Style -Name 'Heading 4' -Size 11 -Color 2f5496 -Italic -Verbose:$false;
                Style -Name 'Heading 5' -Size 11 -Color 2f5496 -Verbose:$false;
                Style -Name 'Heading 6' -Size 11 -Color 1f3763 -Verbose:$false;
                Style -Name TableDefaultHeading -Size 11 -Color fff -Bold -BackgroundColor 4472c4 -Verbose:$false;
                Style -Name TableDefaultRow -Size 11 -Verbose:$false;
                Style -Name TableDefaultAltRow -BackgroundColor d0ddee -Verbose:$false;
                Style -Name Footer -Size 8 -Color 0072af -Hide -Verbose:$false;
                TableStyle TableDefault -BorderWidth 1 -BorderColor 2a70be -HeaderStyle TableDefaultHeading -RowStyle TableDefaultRow -AlternateRowStyle TableDefaultAltRow -Default -Verbose:$false;
                return $pscriboDocument;
            } #end process
        } #end function NewPScriboDocument
        function Invoke-PScriboSection {
        <#
            .SYNOPSIS
                Processes the document/TOC section versioning each level, i.e. 1.2.2.3
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            param ( )
            function Invoke-PScriboSectionLevel {
            <#
                .SYNOPSIS
                    Nested function that processes each document/TOC nested section
            #>
                [CmdletBinding()]
                param (
                    [Parameter(Mandatory)]
                    [ValidateNotNull()]
                    [PSCustomObject] $Section,
                    [Parameter(Mandatory)]
                    [ValidateNotNullOrEmpty()]
                    [System.String] $Number
                )
                if ($pscriboDocument.Options['ForceUppercaseSection']) {
                    $Section.Name = $Section.Name.ToUpper();
                }
                ## Set this section's level
                $Section.Number = $Number;
                $Section.Level = $Number.Split('.').Count -1;
                ### Add to the TOC
                $tocEntry = [PScustomObject] @{ Id = $Section.Id; Number = $Number; Level = $Section.Level; Name = $Section.Name; }
                [ref] $null = $pscriboDocument.TOC.Add($tocEntry);
                ## Set sub-section level seed
                $minorNumber = 1;
                foreach ($s in $Section.Sections) {
                    if ($s.Type -like '*.Section' -and -not $s.IsExcluded) {
                        $sectionNumber = ('{0}.{1}' -f $Number, $minorNumber).TrimStart('.');  ## Calculate section version
                        Invoke-PScriboSectionLevel -Section $s -Number $sectionNumber;
                        $minorNumber++;
                    }
                } #end foreach section
            } #end function Invoke-PScriboSectionLevel
            $majorNumber = 1;
            foreach ($s in $pscriboDocument.Sections) {
                if ($s.Type -like '*.Section') {
                    if ($pscriboDocument.Options['ForceUppercaseSection']) {
                        $s.Name = $s.Name.ToUpper();
                    }
                    if (-not $s.IsExcluded) {
                        Invoke-PScriboSectionLevel -Section $s -Number $majorNumber;
                        $majorNumber++;
                    }
                } #end if
            } #end foreach
        } #end function Invoke-PSScriboSection
        #endregion Document Private Functions
    } #end begin
    process {
        $stopwatch = [Diagnostics.Stopwatch]::StartNew();
        $pscriboDocument = New-PScriboDocument -Name $Name -Id $Id;
        ## Call the Document script block
        foreach ($result in & $ScriptBlock) {
            [ref] $null = $pscriboDocument.Sections.Add($result);
        }
        Invoke-PScriboSection;
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $pscriboDocument.Name);
        $stopwatch.Stop();
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        return $pscriboDocument;
    } #end process
} #end function Document

function DocumentOption {
<#
    .SYNOPSIS
        Initializes a new PScribo global/document options/settings.
    .NOTES
        Options are reset upon each invocation.
#>
    [CmdletBinding(DefaultParameterSetName = 'Margin')]
    [Alias('GlobalOption')]
    param (
        ## Forces document header to be displayed in upper case.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ForceUppercaseHeader,
        ## Forces all section headers to be displayed in upper case.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ForceUppercaseSection,
        ## Enable section/heading numbering
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $EnableSectionNumbering,
        ## Default space replacement separator
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Separator')]
        [AllowNull()]
        [ValidateLength(0,1)]
        [System.String] $SpaceSeparator,
        ## Default page top, bottom, left and right margin (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Margin')]
        [System.UInt16] $Margin = 72,
        ## Default page top and bottom margins (pt)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'CustomMargin')]
        [System.UInt16] $MarginTopAndBottom,
        ## Default page left and right margins (pt)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'CustomMargin')]
        [System.UInt16] $MarginLeftAndRight,
        ## Default page size
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('A4','Legal','Letter')]
        [System.String] $PageSize = 'A4',
        ## Page orientation
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Portrait','Landscape')]
        [System.String] $Orientation = 'Portrait',
        ## Default document font(s)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $DefaultFont = @('Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif')
    )
    process {
        $localized.DocumentOptions | WriteLog;
        if ($SpaceSeparator) {
            WriteLog -Message ($localized.DocumentOptionSpaceSeparator -f $SpaceSeparator);
            $pscriboDocument.Options['SpaceSeparator'] = $SpaceSeparator;
        }
        if ($ForceUppercaseHeader) {
            $localized.DocumentOptionUppercaseHeadings | WriteLog;
            $pscriboDocument.Options['ForceUppercaseHeader'] = $true;
            $pscriboDocument.Name = $pscriboDocument.Name.ToUpper();
        } #end if ForceUppercaseHeader
        if ($ForceUppercaseSection) {
            $localized.DocumentOptionUppercaseSections | WriteLog;
            $pscriboDocument.Options['ForceUppercaseSection'] = $true;
        } #end if ForceUppercaseSection
        if ($EnableSectionNumbering) {
            $localized.DocumentOptionSectionNumbering | WriteLog;
            $pscriboDocument.Options['EnableSectionNumbering'] = $true;
        }
        if ($DefaultFont) {
            WriteLog -Message ($localized.DocumentOptionDefaultFont -f ([System.String]::Join(', ', $DefaultFont)));
            $pscriboDocument.Options['DefaultFont'] = $DefaultFont;
        }
        if ($PSCmdlet.ParameterSetName -eq 'CustomMargin') {
            if ($MarginTopAndBottom -eq 0) { $MarginTopAndBottom = 72; }
            if ($MarginLeftAndRight -eq 0) { $MarginTopAndBottom = 72; }
            $pscriboDocument.Options['MarginTop'] = ConvertPtToMm -Point $MarginTopAndBottom;
            $pscriboDocument.Options['MarginBottom'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginLeft'] = ConvertPtToMm -Point $MarginLeftAndRight;
            $pscriboDocument.Options['MarginRight'] = $pscriboDocument.Options['MarginLeft'];
        }
        else {
            $pscriboDocument.Options['MarginTop'] = ConvertPtToMm -Point $Margin;
            $pscriboDocument.Options['MarginBottom'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginLeft'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginRight'] = $pscriboDocument.Options['MarginTop'];
        }
        WriteLog -Message ($localized.DocumentOptionPageTopMargin -f $pscriboDocument.Options['MarginTop']);
        WriteLog -Message ($localized.DocumentOptionPageRightMargin -f $pscriboDocument.Options['MarginRight']);
        WriteLog -Message ($localized.DocumentOptionPageBottomMargin -f $pscriboDocument.Options['MarginBottom']);
        WriteLog -Message ($localized.DocumentOptionPageLeftMargin -f $pscriboDocument.Options['MarginLeft']);
        ## Convert page size
        ($localized.DocumentOptionPageSize -f $PageSize) | WriteLog;
        switch ($PageSize) {
            'A4' {
                $pscriboDocument.Options['PageWidth'] = 210.0;
                $pscriboDocument.Options['PageHeight'] = 297.0;
            }
            'Legal' {
                $pscriboDocument.Options['PageWidth'] = 215.9;
                $pscriboDocument.Options['PageHeight'] = 355.6;
            }
            'Letter' {
                $pscriboDocument.Options['PageWidth'] = 215.9;
                $pscriboDocument.Options['PageHeight'] = 279.4;
            }
        } #end switch
        ## Convert page size
        ($localized.DocumentOptionPageOrientation -f $Orientation) | WriteLog;
        if ($Orientation -eq 'Landscape') {
            ## Swap the height/width measurements
            $pageHeight = $pscriboDocument.Options['PageHeight'];
            $pscriboDocument.Options['PageHeight'] = $pscriboDocument.Options['PageWidth'];
            $pscriboDocument.Options['PageWidth'] = $pageHeight;
        }
        ($localized.DocumentOptionPageHeight -f $pscriboDocument.Options['PageHeight']) | WriteLog;
        ($localized.DocumentOptionPageWidth -f $pscriboDocument.Options['PageWidth']) | WriteLog;
    } #end process
} #end function DocumentOption

function Export-Document {
<#
    .SYNOPSIS
        Exports a PScribo document object to one or more output formats.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock','')]
    [OutputType([System.IO.FileInfo])]
    param (
        ## PScribo document object
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,
        ## Output formats
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.String[]] $Format,
        ## Output file path
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path = (Get-Location -PSProvider FileSystem),
        ## PScribo document export option
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options,
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $PassThru
    )
    begin {
        try { $Path = Resolve-Path $Path -ErrorAction SilentlyContinue; }
        catch { }
        if ( $(Test-CharsInPath -Path $Path -SkipCheckCharsInFileNamePart) -eq 2 ) {
            throw $localized.IncorrectCharsInPath;
        }
        if (-not (Test-Path $Path -PathType Container)) {
            ## Check $Path is a directory
            throw ($localized.InvalidDirectoryPathError -f $Path);
        }
    }
    process {
        foreach ($f in $Format) {
            WriteLog -Message ($localized.DocumentInvokePlugin -f $f) -Plugin 'Export';
            ## Dynamically generate the output format function name
            $outputFormat = 'Out{0}' -f $f;
            $outputParams = @{
                Document = $Document;
                Path = $Path;
            }
            if ($PSBoundParameters.ContainsKey('Options')) {
                $outputParams['Options'] = $Options;
            }
            $fileInfo = & $outputFormat @outputParams;
            if ($PassThru) {
                Write-Output -InputObject $fileInfo;
            }
        } # end foreach
    } #end process
} #end function Export-Document

function LineBreak {
<#
    .SYNOPSIS
        Initializes a new PScribo line break object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    begin {
        #region LineBreak Private Functions
        function New-PScriboLineBreak {
        <#
            .SYNOPSIS
                Initializes a new PScribo line break object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = [System.Guid]::NewGuid().ToString()
            )
            process {
                $typeName = 'PScribo.LineBreak';
                $pscriboDocument.Properties['LineBreaks']++;
                $pscriboLineBreak = [PSCustomObject] @{
                    Id = $Id;
                    Type = $typeName;
                }
                return $pscriboLineBreak;
            }
        } #end function New-PScriboLineBreak
        #endregion LineBreak Private Functions
    } #end begin
    process {
        WriteLog -Message $localized.ProcessingLineBreak;
        return (New-PScriboLineBreak @PSBoundParameters);
    } #end process
} #end function LineBreak

function PageBreak {
<#
    .SYNOPSIS
        Creates a PScribo page break object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    begin {
        #region PageBreak Private Functions
        function New-PScriboPageBreak {
        <#
            .SYNOPSIS
                Creates a PScribo page break object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = [System.Guid]::NewGuid().ToString()
            )
            process {
                $typeName = 'PScribo.PageBreak';
                $pscriboDocument.Properties['PageBreaks']++;
                $pscriboPageBreak = [PSCustomObject] @{
                    Id = $Id;
                    Type = $typeName;
                }
                return $pscriboPageBreak;
            }
        } #end function New-PScriboPageBreak
        #endregion PageBreak Private Functions
    } #end begin
    process {
        WriteLog -Message $localized.ProcessingPageBreak;
        return (New-PScriboPageBreak -Id $Id);
    }
} #end function PageBreak

function Paragraph {
<#
    .SYNOPSIS
        Initializes a new PScribo paragraph object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## Paragraph Id and Xml element name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,
        ## Paragraph text. If empty $Name/Id will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [AllowNull()]
        [System.String] $Text = $null,
        ## Output value override, i.e. for Xml elements. If empty $Text will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [AllowNull()]
        [System.String] $Value = $null,
        ## Paragraph style Name/Id reference.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Style = $null,
        ## No new line - ONLY IMPLEMENTED FOR TEXT OUTPUT
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoNewLine,
        ## Override the bold style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Bold,
        ## Override the italic style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Italic,
        ## Override the underline style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Underline,
        ## Override the font name(s)
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Font,
        ## Override the font size (pt)
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.UInt16] $Size = $null,
        ## Override the font color/colour
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Color = $null,
        ## Tab indent
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0
    )
    begin {
        #region Paragraph Private Functions
        function New-PScriboParagraph {
        <#
            .SYNOPSIS
                Initializes a new PScribo paragraph object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## Paragraph Id (and Xml) element name
                [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,
                ## Paragraph text. If empty $Name/Id will be used.
                [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
                [AllowNull()]
                [System.String] $Text = $null,
                ## Ouptut value override, i.e. for Xml elements. If empty $Text will be used.
                [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
                [AllowNull()]
                [System.String] $Value = $null,
                ## Paragraph style Name/Id reference.
                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.String] $Style = $null,
                ## No new line - ONLY IMPLEMENTED FOR TEXT OUTPUT
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $NoNewLine,
                ## Override the bold style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Bold,
                ## Override the italic style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Italic,
                ## Override the underline style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Underline,
                ## Override the font name(s)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String[]] $Font,
                ## Override the font size (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.UInt16] $Size = $null,
                ## Override the font color/colour
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('Colour')]
                [AllowNull()]
                [System.String] $Color = $null,
                ## Tab indent
                [Parameter()]
                [ValidateRange(0,10)]
                [System.Int32] $Tabs = 0
            )
            begin {
                if (-not ([string]::IsNullOrEmpty($Text))) {
                    $Name = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                }
                if ($Color) {
                    $Color = Resolve-PScriboStyleColor -Color $Color;
                }
            } #end begin
            process {
                $typeName = 'PScribo.Paragraph';
                $pscriboDocument.Properties['Paragraphs']++;
                $pscriboParagraph = [PSCustomObject] @{
                    Id = $Name;
                    Text = $Text;
                    Type = $typeName;
                    Style = $Style;
                    Value = $Value;
                    NewLine = !$NoNewLine;
                    Tabs = $Tabs;
                    Bold = $Bold;
                    Italic = $Italic;
                    Underline = $Underline;
                    Font = $Font;
                    Size = $Size;
                    Color = $Color;
                }
                return $pscriboParagraph;
            } #end process
        } #end function New-PScriboParagraph
        #endregion Paragraph Private Functions
    } #end begin
    process {
        if ($Name.Length -gt 40) { $paragraphDisplayName = '{0}[..]' -f $Name.Substring(0,36); }
        else { $paragraphDisplayName = $Name; }
        WriteLog -Message ($localized.ProcessingParagraph -f $paragraphDisplayName);
        return (New-PScriboParagraph @PSBoundParameters);
    } #end process
} #end function Paragraph

function Section {
<#
    .SYNOPSIS
        Initializes a new PScribo section object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## PScribo section heading/name.
        [Parameter(Mandatory, Position = 0)]
        [System.String] $Name,
        ## PScribo document script block.
        [Parameter(Position = 1)]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),
        ## PScribo style applied to document section.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Style = $null,
        ## Section is excluded from TOC/section numbering.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ExcludeFromTOC
    )
    begin {
        #region Section Private Functions
        function New-PScriboSection {
        <#
            .SYNOPSIS
                Initializes new PScribo section object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## PScribo section heading/name.
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,
                ## PScribo style applied to document section.
                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.String] $Style = $null,
                ## Section is excluded from TOC/section numbering.
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $IsExcluded
            )
            process {
                $typeName = 'PScribo.Section';
                $pscriboDocument.Properties['Sections']++;
                $pscriboSection = [PSCustomObject] @{
                    Id = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                    Level = 0;
                    Number = '';
                    Name = $Name;
                    Type = $typeName;
                    Style = $Style;
                    IsExcluded = $IsExcluded;
                    Sections = (New-Object -TypeName System.Collections.ArrayList);
                }
                return $pscriboSection;
            } #end process
        } #end function new-pscribosection
        #endregion Section Private Functions
    } #end begin
    process {
        WriteLog -Message ($localized.ProcessingSectionStarted -f $Name);
        $pscriboSection = New-PScriboSection -Name $Name -Style $Style -IsExcluded:$ExcludeFromTOC;
        foreach ($result in & $ScriptBlock) {
            ## Ensure we don't have something errant passed down the pipeline (#29)
            if ($result -is [System.Management.Automation.PSCustomObject]) {
                if (('Id' -in $result.PSObject.Properties.Name) -and
                    ('Type' -in $result.PSObject.Properties.Name) -and
                    ($result.Type -match '^PScribo.')) {
                    [ref] $null = $pscriboSection.Sections.Add($result);
                }
                else {
                    WriteLog -Message ($localized.UnexpectedObjectWarning -f $Name) -IsWarning;
                }
            }
            else {
                WriteLog -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), $Name) -IsWarning;
            }
        }
        WriteLog -Message ($localized.ProcessingSectionCompleted -f $Name);
        return $pscriboSection;
    } #end process
} #end function Section

function Resolve-PScriboStyleColor {
<#
    .SYNOPSIS
        Resolves a HTML color format or Word color constant to a RGB value
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNull()]
        [System.String] $Color
    )
    begin {
        # http://www.jadecat.com/tuts/colorsplus.html
        $wordColorConstants = @{
            AliceBlue = 'F0F8FF'; AntiqueWhite = 'FAEBD7'; Aqua = '00FFFF'; Aquamarine = '7FFFD4'; Azure = 'F0FFFF'; Beige = 'F5F5DC';
            Bisque = 'FFE4C4'; Black = '000000'; BlanchedAlmond = 'FFEBCD'; Blue = '0000FF'; BlueViolet = '8A2BE2'; Brown = 'A52A2A';
            BurlyWood = 'DEB887'; CadetBlue = '5F9EA0'; Chartreuse = '7FFF00'; Chocolate = 'D2691E'; Coral = 'FF7F50';
            CornflowerBlue = '6495ED'; Cornsilk = 'FFF8DC'; Crimson = 'DC143C'; Cyan = '00FFFF'; DarkBlue = '00008B'; DarkCyan = '008B8B';
            DarkGoldenrod = 'B8860B'; DarkGray = 'A9A9A9'; DarkGreen = '006400'; DarkKhaki = 'BDB76B'; DarkMagenta = '8B008B';
            DarkOliveGreen = '556B2F'; DarkOrange = 'FF8C00'; DarkOrchid = '9932CC'; DarkRed = '8B0000'; DarkSalmon = 'E9967A';
            DarkSeaGreen = '8FBC8F'; DarkSlateBlue = '483D8B'; DarkSlateGray = '2F4F4F'; DarkTurquoise = '00CED1'; DarkViolet = '9400D3';
            DeepPink = 'FF1493'; DeepSkyBlue = '00BFFF'; DimGray = '696969'; DodgerBlue = '1E90FF'; Firebrick = 'B22222';
            FloralWhite = 'FFFAF0'; ForestGreen = '228B22'; Fuchsia = 'FF00FF'; Gainsboro = 'DCDCDC'; GhostWhite = 'F8F8FF';
            Gold = 'FFD700'; Goldenrod = 'DAA520'; Gray = '808080'; Green = '008000'; GreenYellow = 'ADFF2F'; Honeydew = 'F0FFF0';
            HotPink = 'FF69B4'; IndianRed = 'CD5C5C'; Indigo = '4B0082'; Ivory = 'FFFFF0'; Khaki = 'F0E68C'; Lavender = 'E6E6FA';
            LavenderBlush = 'FFF0F5'; LawnGreen = '7CFC00'; LemonChiffon = 'FFFACD'; LightBlue = 'ADD8E6'; LightCoral = 'F08080';
            LightCyan = 'E0FFFF'; LightGoldenrodYellow = 'FAFAD2'; LightGreen = '90EE90'; LightGrey = 'D3D3D3'; LightPink = 'FFB6C1';
            LightSalmon = 'FFA07A'; LightSeaGreen = '20B2AA'; LightSkyBlue = '87CEFA'; LightSlateGray = '778899'; LightSteelBlue = 'B0C4DE';
            LightYellow = 'FFFFE0'; Lime = '00FF00'; LimeGreen = '32CD32'; Linen = 'FAF0E6'; Magenta = 'FF00FF'; Maroon = '800000';
            McMintGreen = 'BED6C9'; MediumAuqamarine = '66CDAA'; MediumBlue = '0000CD'; MediumOrchid = 'BA55D3'; MediumPurple = '9370D8';
            MediumSeaGreen = '3CB371'; MediumSlateBlue = '7B68EE'; MediumSpringGreen = '00FA9A'; MediumTurquoise = '48D1CC';
            MediumVioletRed = 'C71585'; MidnightBlue = '191970'; MintCream = 'F5FFFA'; MistyRose = 'FFE4E1'; Moccasin = 'FFE4B5';
            NavajoWhite = 'FFDEAD'; Navy = '000080'; OldLace = 'FDF5E6'; Olive = '808000'; OliveDrab = '688E23'; Orange = 'FFA500';
            OrangeRed = 'FF4500'; Orchid = 'DA70D6'; PaleGoldenRod = 'EEE8AA'; PaleGreen = '98FB98'; PaleTurquoise = 'AFEEEE';
            PaleVioletRed = 'D87093'; PapayaWhip = 'FFEFD5'; PeachPuff = 'FFDAB9'; Peru = 'CD853F'; Pink = 'FFC0CB'; Plum = 'DDA0DD';
            PowderBlue = 'B0E0E6'; Purple = '800080'; Red = 'FF0000'; RosyBrown = 'BC8F8F'; RoyalBlue = '4169E1'; SaddleBrown = '8B4513';
            Salmon = 'FA8072'; SandyBrown = 'F4A460'; SeaGreen = '2E8B57'; Seashell = 'FFF5EE'; Sienna = 'A0522D'; Silver = 'C0C0C0';
            SkyBlue = '87CEEB'; SlateBlue = '6A5ACD'; SlateGray = '708090'; Snow = 'FFFAFA'; SpringGreen = '00FF7F'; SteelBlue = '4682B4';
            Tan = 'D2B48C'; Teal = '008080'; Thistle = 'D8BFD8'; Tomato = 'FF6347'; Turquoise = '40E0D0'; Violet = 'EE82EE'; Wheat = 'F5DEB3';
            White = 'FFFFFF'; WhiteSmoke = 'F5F5F5'; Yellow = 'FFFF00'; YellowGreen = '9ACD32';
        };
    } #end begin
    process {
        $pscriboColor = $Color;
        if ($wordColorConstants.ContainsKey($pscriboColor)) {
            return $wordColorConstants[$pscriboColor].ToLower();
        }
        elseif ($pscriboColor.Length -eq 6 -or $pscriboColor.Length -eq 3) {
            $pscriboColor = '#{0}' -f $pscriboColor;
        }
        elseif ($pscriboColor.Length -eq 7 -or $pscriboColor.Length -eq 4) {
            if (-not ($pscriboColor.StartsWith('#'))) {
                return $null;
            }
        }
        if ($pscriboColor -notmatch '^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$') {
            return $null;
        }
        return $pscriboColor.TrimStart('#').ToLower();
    } #end process
} #end function ResolvePScriboColor
function Test-PScriboStyleColor {
<#
    .SYNOPSIS
        Tests whether a color string is a valid HTML color.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color
    )
    process {
        if (Resolve-PScriboStyleColor -Color $Color) { return $true; }
        else { return $false; }
    } #end process
} #end function test-pscribostylecolor
function Test-PScriboStyle {
<#
    .SYNOPSIS
        Tests whether a style has been defined.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name
    )
    process {
        return $PScriboDocument.Styles.ContainsKey($Name);
    }
} #end function Test-PScriboStyle
function Style {
<#
    .SYNOPSIS
        Defines a new PScribo formatting style.
    .DESCRIPTION
        Creates a standard format formatting style that can be applied
        to PScribo document keywords, e.g. a combination of font style, font
        weight and font size.
    .NOTES
        Not all plugins support all options.
#>
    [CmdletBinding()]
    param (
        ## Style name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,
        ## Font size (pt)
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [System.UInt16] $Size = 11,
        ## Font color/colour
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Colour')]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color = '000',
        ## Background color/colour
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('BackgroundColour')]
        [ValidateNotNullOrEmpty()]
        [System.String] $BackgroundColor,
        ## Bold typeface
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Bold,
        ## Italic typeface
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Italic,
        ## Underline typeface
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Underline,
        ## Text alignment
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Left','Center','Right','Justify')]
        [System.String] $Align = 'Left',
        ## Set as default style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Default,
        ## Style id
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = $Name -Replace(' ',''),
        ## Font name (array of names for HTML output)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $Font,
        ## Html CSS class id - to override Style.Id in HTML output.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $ClassId = $Id,
        ## Hide style from UI (Word)
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Hide')]
        [System.Management.Automation.SwitchParameter] $Hidden
    )
    begin {
        #region Style Private Functions
        function Add-PScriboStyle {
        <#
            .SYNOPSIS
                Initializes a new PScribo style object.
        #>
            [CmdletBinding()]
            param (
                ## Style name
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,
                ## Style id
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = $Name -Replace(' ',''),
                ## Font size (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.UInt16] $Size = 11,
                ## Font name (array of names for HTML output)
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.String[]] $Font,
                ## Font color/colour
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('Colour')]
                [ValidateNotNullOrEmpty()]
                [System.String] $Color = 'Black',
                ## Background color/colour
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('BackgroundColour')]
                [ValidateNotNullOrEmpty()]
                [System.String] $BackgroundColor,
                ## Bold typeface
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Bold,
                ## Italic typeface
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Italic,
                ## Underline typeface
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Underline,
                ## Text alignment
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('Left','Center','Right','Justify')]
                [string] $Align = 'Left',
                ## Html CSS class id - to override Style.Id in HTML output.
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.String] $ClassId = $Id,
                ## Hide style from UI (Word)
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('Hide')]
                [System.Management.Automation.SwitchParameter] $Hidden,
                ## Set as default style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Default
            ) #end param
            begin {
                if (-not (Test-PScriboStyleColor -Color $Color)) {
                    throw ($localized.InvalidHtmlColorError -f $Color);
                }
                if ($BackgroundColor) {
                    if (-not (Test-PScriboStyleColor -Color $BackgroundColor)) {
                        throw ($localized.InvalidHtmlBackgroundColorError -f $BackgroundColor);
                    }
                    else {
                        $BackgroundColor = Resolve-PScriboStyleColor -Color $BackgroundColor;
                    }
                }
                if (-not ($Font)) {
                    $Font = $pscriboDocument.Options['DefaultFont'];
                }
            } #end begin
            process {
                $pscriboDocument.Properties['Styles']++;
                $style = [PSCustomObject] @{
                    Id   = $Id;
                    Name = $Name;
                    Font = $Font;
                    Size = $Size;
                    Color = (Resolve-PScriboStyleColor -Color $Color).ToLower();
                    BackgroundColor = $BackgroundColor.ToLower();
                    Bold = $Bold.ToBool();
                    Italic = $Italic.ToBool();
                    Underline = $Underline.ToBool();
                    Align = $Align;
                    ClassId = $ClassId;
                    Hidden = $Hidden.ToBool();
                }
                $pscriboDocument.Styles[$Id] = $style;
                if ($Default) { $pscriboDocument.DefaultStyle = $style.Id; }
            } #end process
        } #end function Add-PScriboStyle
        #endregion Style Private Functions
    }
    process {
        WriteLog -Message ($localized.ProcessingStyle -f $Id);
        Add-PScriboStyle @PSBoundParameters;
    } #end process
} #end function Style
function Set-Style {
<#
    .SYNOPSIS
        Sets the style for an individual table row or cell.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Object])]
    param (
        ## PSCustomObject to apply the style to
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object[]] [Ref] $InputObject,
        ## PScribo style Id to apply
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $Style,
        ## Property name(s) to apply the selected style to. Leave blank to apply the style to the entire row.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Property = '',
        ## Passes the modified object back to the pipeline
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $PassThru
    ) #end param
    begin {
        if (-not (Test-PScriboStyle -Name $Style)) {
            Write-Error ($localized.UndefinedStyleError -f $Style);
            return;
        }
    }
    process {
        foreach ($object in $InputObject) {
            foreach ($p in $Property) {
                ## If $Property not set, __Style will apply to the whole row.
                $propertyName = '{0}__Style' -f $p;
                $object | Add-Member -MemberType NoteProperty -Name $propertyName -Value $Style -Force;
            }
        }
        if ($PassThru) {
            return $object;
        }
    } #end process
} #end function Set-Style

function Table {
<#
    .SYNOPSIS
        Defines a new PScribo document table.
    .PARAMETER Name
    .PARAMETER InputObject
    .PARAMETER Hashtable
    .PARAMETER Columns
    .PARAMETER ColumnWidths
    .PARAMETER Headers
    .PARAMETER Style
    .PARAMETER List
    .PARAMETER Width
    .PARAMETER Tabs
    .EXAMPLE
        Table -Name 'Table 1' -InputObject $(Get-Service) -Columns 'Name','DisplayName','Status' -ColumnWidths 40,20,40
#>
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## Table name/Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string] $Name = ([System.Guid]::NewGuid().ToString()),
        # Array of objects
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
        [Alias('CustomObject','Object')]
        [ValidateNotNullOrEmpty()]
        [System.Object[]] $InputObject,
        # Array of Hashtables
        [Parameter(Mandatory, ParameterSetName = 'Hashtable')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Specialized.OrderedDictionary[]] $Hashtable,
        # Array of Hashtable key names or Object/PSCustomObject property names to include, in display order.
        # If not supplied then all Hashtable keys or all PSCustomObject properties will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'InputObject')]
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'Hashtable')]
        [Alias('Properties')]
        [AllowNull()]
        [System.String[]] $Columns = $null,
        ## Column widths as percentages. Total should not exceed 100.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.UInt16[]] $ColumnWidths,
        # Array of custom table header strings in display order.
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [AllowNull()]
        [System.String[]] $Headers = $null,
        ## Table style
        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Style = $pscriboDocument.DefaultTableStyle,
        # List view (no headers)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $List,
        ## Table width (%), 0 = Autofit
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,100)]
        [System.UInt16] $Width = 100,
        ## Indent table
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.UInt16] $Tabs
    ) #end param
    begin {
        #region Table Private Functions
        function New-PScriboTable {
        <#
            .SYNOPSIS
                Initializes a new PScribo table object.
            .PARAMETER Name
            .PARAMETER Columns
            .PARAMETER ColumnWidths
            .PARAMETER Rows
            .PARAMETER Style
            .PARAMETER Width
            .PARAMETER List
            .PARAMETER Width
            .PARAMETER Tabs
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## Table name/Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name = ([System.Guid]::NewGuid().ToString()),
                ## Table columns/display order
                [Parameter(Mandatory)]
                [AllowNull()]
                [System.String[]] $Columns,
                ## Table columns widths
                [Parameter(Mandatory)]
                [AllowNull()]
                [System.UInt16[]] $ColumnWidths,
                ## Collection of PScriboTableObjects for table rows
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                [System.Collections.ArrayList] $Rows,
                ## Table style
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Style,
                ## List view
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $List,
                ## Table width (%), 0 = Autofit
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateRange(0,100)]
                [System.UInt16] $Width = 100,
                ## Indent table
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateRange(0,10)]
                [System.UInt16] $Tabs
            ) #end param
            process {
                $typeName = 'PScribo.Table';
                $pscriboDocument.Properties['Tables']++;
                $pscriboTable = [PSCustomObject] @{
                    Id = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                    Name = $Name;
                    Type = $typeName;
                    # Headers = $Headers; ## Headers are stored as they may be required when formatting output, i.e. Word tables
                    Columns = $Columns;
                    ColumnWidths = $ColumnWidths;
                    Rows = $Rows;
                    List = $List;
                    Style = $Style;
                    Width = $Width;
                    Tabs = $Tabs;
                }
                return $pscriboTable;
            } #end process
        } #end function new-pscribotable
        function New-PScriboTableRow {
        <#
            .SYNOPSIS
                Defines a new PScribo document table row from an object or hashtable.
            .PARAMETER InputObject
            .PARAMETER Properties
            .PARAMETER Headers
            .PARAMETER Hashtable
        #>
            [CmdletBinding(DefaultParameterSetName = 'InputObject')]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## PSCustomObject to create PScribo table row
                [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
                [ValidateNotNull()]
                [System.Object] $InputObject,
                ## PSCutomObject properties to include in the table row
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'InputObject')]
                [AllowNull()]
                [System.String[]] $Properties,
                # Custom table header strings (in Display Order). Used for property names.
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'InputObject')]
                [AllowNull()]
                [System.String[]] $Headers = $null,
                ## Array of ordered dictionaries (hashtables) to create PScribo table row
                [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Hashtable')]
                [AllowNull()]
                [System.Collections.Specialized.OrderedDictionary] $Hashtable
            )
            begin {
                Write-Debug ('Using parameter set "{0}.' -f $PSCmdlet.ParameterSetName);
            } #end begin
            process {
                switch ($PSCmdlet.ParameterSetName) {
                    'Hashtable'{
                        if (-not $Hashtable.Contains('__Style')) {
                            $Hashtable['__Style'] = $null;
                        }
                        ## Create and return custom object from hashtable
                        return ([PSCustomObject] $Hashtable);
                    } #end Hashtable
                    Default {
                        $objectProperties = [Ordered] @{ };
                        if ($Properties -notcontains '__Style') { $Properties += '__Style'; }
                        ## Build up hashtable of required property names
                        for ($i = 0; $i -lt $Properties.Count; $i++) {
                            $propertyName = $Properties[$i];
                            $propertyStyleName = '{0}__Style' -f $propertyName;
                            if ($InputObject.PSObject.Properties[$propertyStyleName]) {
                                if ($Headers) {
                                    ## Rename the style property to match the header
                                    $headerStyleName = '{0}__Style' -f $Headers[$i];
                                    $objectProperties[$headerStyleName] = $InputObject.$propertyStyleName;
                                }
                                else {
                                    $objectProperties[$propertyStyleName] = $InputObject.$propertyStyleName;
                                }
                            }
                            if ($Headers -and $PropertyName -notlike '*__Style') {
                                if ($InputObject.PSObject.Properties[$propertyName]) {
                                    $objectProperties[$Headers[$i]] = $InputObject.$propertyName;
                                }
                            }
                            else {
                                if ($InputObject.PSObject.Properties[$propertyName]) {
                                    $objectProperties[$propertyName] = $InputObject.$propertyName;
                                }
                                else {
                                    $objectProperties[$propertyName] = $null;
                                }
                            }
                        } #end for
                        ## Create and return custom object
                        return ([PSCustomObject] $objectProperties);
                    } #end Default
                } #end switch
            } #end process
        } #end function New-PScriboTableRow
        #endregion Table Private Functions
        Write-Debug ('Using parameter set "{0}".' -f $PSCmdlet.ParameterSetName);
        [System.Collections.ArrayList] $rows = New-Object -TypeName System.Collections.ArrayList;
        WriteLog -Message ($localized.ProcessingTable -f $Name);
        if ($Headers -and (-not $Columns)) {
            WriteLog -Message $localized.TableHeadersWithNoColumnsWarning -IsWarning;
            $Headers = $Columns;
        } #end if
        elseif (($null -ne $Columns) -and ($null -ne $Headers)) {
            ## Check the number of -Headers matches the number of -Properties
            if ($Headers.Count -ne $Columns.Count) {
                WriteLog -Message $localized.TableHeadersCountMismatchWarning -IsWarning;
                $Headers = $Columns;
            }
        } #end if
        if ($ColumnWidths) {
            $columnWidthsSum = $ColumnWidths | Measure-Object -Sum | Select-Object -ExpandProperty Sum;
            if ($columnWidthsSum -ne 100) {
                WriteLog -Message ($localized.TableColumnWidthSumWarning -f $columnWidthsSum) -IsWarning;
                $ColumnWidths = $null;
            }
            elseif ($List -and $ColumnWidths.Count -ne 2) {
                WriteLog -Message $localized.ListTableColumnCountWarning -IsWarning;
                $ColumnWidths = $null;
            }
            elseif (($PSCmdlet.ParameterSetName -eq 'Hashtable') -and (-not $List) -and ($Hashtable[0].Keys.Count -ne $ColumnWidths.Count)) {
                WriteLog -Message $localized.TableColumnWidthMismatchWarning -IsWarning;
                $ColumnWidths = $null;
            }
            elseif (($PSCmdlet.ParameterSetName -eq 'InputObject') -and (-not $List)) {
                ## Columns might not have been passed and there is no object in the pipeline here, so check $Columns is an array.
                if (($Columns -is [System.Object[]]) -and ($Columns.Count -ne $ColumnWidths.Count)) {
                    WriteLog -Message $localized.TableColumnWidthMismatchWarning -IsWarning;
                    $ColumnWidths = $null;
                }
            }
        } #end if columnwidths
    } #end begin
    process {
        if ($null -eq $Columns) {
            ## Use all available properties
            switch ($PSCmdlet.ParameterSetName) {
                'Hashtable' {
                    $Columns = $Hashtable | Select-Object -First 1 -ExpandProperty Keys | Where-Object { $_ -notlike '*__Style' };
                }
                Default {
                    ## Pipeline objects are not available in the begin scriptblock
                    $object = $InputObject | Select-Object -First 1;
                    if ($object -is [System.Management.Automation.PSCustomObject]) {
                        $Columns = $object.PSObject.Properties | Where-Object Name -notlike '*__Style' | Select-Object -ExpandProperty Name;
                    }
                    else {
                        $Columns = Get-Member -InputObject $object -MemberType Properties | Where-Object Name -notlike '*__Style' | Select-Object -ExpandProperty Name;
                    }
                } #end default
            } #end switch parametersetname
        } # end if not columns
        switch ($PSCmdlet.ParameterSetName) {
            'Hashtable' {
                foreach ($nestedHashtable in $Hashtable) {
                    $customObject = New-PScriboTableRow -Hashtable $nestedHashtable;
                    [ref] $null = $rows.Add($customObject);
                } #end foreach nested hashtable entry
            } #end hashtable
            Default {
                foreach ($object in $InputObject) {
                    $customObject = New-PScriboTableRow -InputObject $object -Properties $Columns -Headers $Headers;
                    [ref] $null = $rows.Add($customObject);
                } #end foreach inputobject
            } #end default
        } #end switch
    } #end process
    end {
        ## Reset the column names as the object have been rewritten with their headers
        if ($Headers) { $Columns = $Headers; }
        $table = @{
            Name = $Name;
            Columns = $Columns;
            ColumnWidths = $ColumnWidths;
            Rows = $rows;
            List = $List;
            Style = $Style;
            Width = $Width;
            Tabs = $Tabs;
        }
        return (New-PScriboTable @table);
    } #end end
} #end function Table

function TableStyle {
<#
    .SYNOPSIS
        Defines a new PScribo table formatting style.
    .DESCRIPTION
        Creates a standard table formatting style that can be applied
        to the PScribo table keyword, e.g. a combination of header and
        row styles and borders.
    .NOTES
        Not all plugins support all options.
#>
    [CmdletBinding()]
    param (
        ## Table Style name/id
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String] $Id,
        ## Header Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String] $HeaderStyle = 'Default',
        ## Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String] $RowStyle = 'Default',
        ## Header Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [AllowNull()]
        [Alias('AlternatingRowStyle')]
        [System.String] $AlternateRowStyle = 'Default',
        ## Table border size/width (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Border')]
        [AllowNull()]
        [System.Single] $BorderWidth = 0,
        ## Table border colour
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Border')]
        [ValidateNotNullOrEmpty()]
        [Alias('BorderColour')]
        [System.String] $BorderColor = '000',
        ## Table cell top padding (pt)
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Single] $PaddingTop = 1.0,
        ## Table cell left padding (pt)
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Single] $PaddingLeft = 4.0,
        ## Table cell bottom padding (pt)
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Single] $PaddingBottom = 0.0,
        ## Table cell right padding (pt)
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Single] $PaddingRight = 4.0,
        ## Table alignment
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Left','Center','Right')]
        [System.String] $Align = 'Left',
        ## Set as default table style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Default
    ) #end param
    begin {
        #region TableStyle Private Functions
        function Add-PScriboTableStyle {
        <#
            .SYNOPSIS
                Defines a new PScribo table formatting style.
            .DESCRIPTION
                Creates a standard table formatting style that can be applied
                to the PScribo table keyword, e.g. a combination of header and
                row styles and borders.
            .NOTES
                Not all plugins support all options.
        #>
            [CmdletBinding()]
            param (
                ## Table Style name/id
                [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
                [ValidateNotNullOrEmpty()]
                [Alias('Name')]
                [System.String] $Id,
                ## Header Row Style Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
                [ValidateNotNullOrEmpty()]
                [System.String] $HeaderStyle = 'Normal',
                ## Row Style Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
                [ValidateNotNullOrEmpty()]
                [System.String] $RowStyle = 'Normal',
                ## Header Row Style Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
                [AllowNull()]
                [Alias('AlternatingRowStyle')]
                [System.String] $AlternateRowStyle = 'Normal',
                ## Table border size/width (pt)
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Border')]
                [AllowNull()]
                [System.Single] $BorderWidth = 0,
                ## Table border colour
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Border')]
                [ValidateNotNullOrEmpty()]
                [Alias('BorderColour')]
                [System.String] $BorderColor = '000',
                ## Table cell top padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingTop = 1.0,
                ## Table cell left padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingLeft = 4.0,
                ## Table cell bottom padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingBottom = 0.0,
                ## Table cell right padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingRight = 4.0,
                ## Table alignment
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('Left','Center','Right')]
                [System.String] $Align = 'Left',
                ## Set as default table style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Default
            ) #end param
            begin {
                if ($BorderWidth -gt 0) { $borderStyle = 'Solid'; } else {$borderStyle = 'None'; }
                if (-not ($pscriboDocument.Styles.ContainsKey($HeaderStyle))) {
                    throw ($localized.UndefinedTableHeaderStyleError -f $HeaderStyle);
                }
                if (-not ($pscriboDocument.Styles.ContainsKey($RowStyle))) {
                    throw ($localized.UndefinedTableRowStyleError -f $RowStyle);
                }
                if (-not ($pscriboDocument.Styles.ContainsKey($AlternateRowStyle))) {
                    throw ($localized.UndefinedAltTableRowStyleError -f $AlternateRowStyle);
                }
                if (-not (Test-PScriboStyleColor -Color $BorderColor)) {
                    throw ($localized.InvalidTableBorderColorError -f $BorderColor);
                }
            } #end begin
            process {
                $pscriboDocument.Properties['TableStyles']++;
                $tableStyle = [PSCustomObject] @{
                    Id = $Id.Replace(' ', $pscriboDocument.Options['SpaceSeparator']);
                    Name = $Id;
                    HeaderStyle = $HeaderStyle;
                    RowStyle = $RowStyle;
                    AlternateRowStyle = $AlternateRowStyle;
                    PaddingTop = ConvertPtToMm $PaddingTop;
                    PaddingLeft = ConvertPtToMm $PaddingLeft;
                    PaddingBottom = ConvertPtToMm $PaddingBottom;
                    PaddingRight = ConvertPtToMm $PaddingRight;
                    Align = $Align;
                    BorderWidth = ConvertPtToMm $BorderWidth;
                    BorderStyle = $borderStyle;
                    BorderColor = Resolve-PScriboStyleColor -Color $BorderColor;
                }
                $pscriboDocument.TableStyles[$Id] = $tableStyle;
                if ($Default) { $pscriboDocument.DefaultTableStyle = $tableStyle.Id; }
            } #end process
        } #end function Add-PScriboTableStyle
        #endregion TableStyle Private Functions
    }
    process {
        WriteLog -Message ($localized.ProcessingTableStyle -f $Id);
        Add-PScriboTableStyle @PSBoundParameters;
    }
} #end function tablestyle

function TOC {
<#
    .SYNOPSIS
        Initializes a new PScribo Table of Contents (TOC) object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        [Parameter(ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name = 'Contents',
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $ClassId = 'TOC'
    )
    begin {
        #region TOC Private Functions
        function New-PScriboTOC {
        <#
            .SYNOPSIS
                Initializes a new PScribo Table of Contents (TOC) object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(ValueFromPipeline)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name = 'Contents',
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String] $ClassId = 'TOC'
            )
            process {
                $typeName = 'PScribo.TOC';
                if ($pscriboDocument.Options['ForceUppercaseSection']) {
                    $Name = $Name.ToUpper();
                }
                $pscriboDocument.Properties['TOCs']++;
                $pscriboTOC = [PSCustomObject] @{
                    Id = [System.Guid]::NewGuid().ToString();
                    Name = $Name;
                    Type = $typeName;
                    ClassId = $ClassId;
                }
                return $pscriboTOC;
            } #end process
        } #end function New-PScriboTOC
        #endregion TOC Private Functions
    } #end begin
    process {
        WriteLog -Message ($localized.ProcessingTOC -f $Name);
        return (New-PScriboTOC @PSBoundParameters);
    }
} #end function TOC

function ConvertPtToMm {
<#
    .SYNOPSIS
        Convert points into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('pt')]
        [System.Single] $Point
    )
    process {
        return [System.Math]::Round(($Point / 72) * 25.4, 2);
    }
} #end function ConvertPtToMm
function ConvertPxToMm {
<#
    .SYNOPSIS
        Convert pixels into millimeters (default 96dpi)
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('px')]
        [System.Single] $Pixel,
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int16] $Dpi = 96
    )
    process {
        return [System.Math]::Round((25.4 / $Dpi) * $Pixel, 2);
    }
} #end function ConvertPxToMm
function ConvertInToMm {
<#
    .SYNOPSIS
        Convert inches into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('in')]
        [System.Single] $Inch
    )
    process {
        return [System.Math]::Round($Inch * 25.4, 2);
    }
} #end function ConvertInToMm
function ConvertMmToIn {
<#
    .SYNOPSIS
        Convert millimeters into inches
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {
        return [System.Math]::Round($Millimeter / 25.4, 2);
    }
} #end function ConvertMmToIn
function ConvertMmToPt {
<#
    .SYNOPSIS
        Convert millimeters into points
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {
        return ((ConvertMmToIn $Millimeter) / 0.0138888888888889);
    }
} #end function ConvertMmToPt
function ConvertMmToTwips {
<#
    .SYNOPSIS
        Convert millimeters into twips
    .NOTES
        1 twip = 1/20th pt
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {
        return (ConvertMmToIn -Millimeter $Millimeter) * 1440;
    }
} #end function ConvertMmToTwips
function ConvertMmToOctips {
<#
    .SYNOPSIS
        Convert millimeters into octips
    .NOTES
        1 "octip" = 1/8th pt
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {
        return (ConvertMmToIn -Millimeter $Millimeter) * 576;
    }
} #end function ConvertMmToOctips
function ConvertMmToEm {
<#
    .SYNOPSIS
        Convert millimeters into em
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {
        return [System.Math]::Round($Millimeter / 4.23333333333333, 2);
    }
} #end function ConvertMmToEm
function ConvertMmToPx {
<#
    .SYNOPSIS
        Convert millimeters into pixels (default 96dpi)
#>
    [CmdletBinding()]
    [OutputType([System.Int16])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter,
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int16] $Dpi = 96
    )
    process {
        $pixels = [System.Int16] ((ConvertMmToIn -Millimeter $Millimeter) * $Dpi);
        if ($pixels -lt 1) { return (1 -as [System.Int16]); }
        else { return $pixels; }
    }
} #end function ConvertMmToPx
function ConvertToInvariantCultureString {
    <#
        .SYNOPSIS
            Convert to a number to a string with a culture-neutral representation #6, #42.
    #>
    [CmdletBinding()]
    param (
        ## The sinle/double
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Object] $Object,
        ## Format string
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Format
    )
    if ($PSBoundParameters.ContainsKey('Format')) {
        return $Object.ToString($Format, [System.Globalization.CultureInfo]::InvariantCulture);
    }
    else {
        return $Object.ToString([System.Globalization.CultureInfo]::InvariantCulture);
    }
} #end function ConvertToInvariantCultureString

function WriteLog {
<#
    .SYNOPSIS
        Writes message to the verbose, warning or debug streams. Output is
        prefixed with the time and PScribo plugin name.
#>
    [CmdletBinding(DefaultParameterSetName = 'Verbose')]
    param (
        ## Message to send to the Verbose stream
        [Parameter(ValueFromPipeline, ParameterSetName = 'Verbose')]
        [Parameter(ValueFromPipeline, ParameterSetName = 'Warning')]
        [Parameter(ValueFromPipeline, ParameterSetName = 'Debug')]
        [ValidateNotNullOrEmpty()]
        [System.String] $Message,
        ## PScribo plugin name
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Plugin,
        ## Redirect message to the Warning stream
        [Parameter(ParameterSetName = 'Warning')]
        [System.Management.Automation.SwitchParameter] $IsWarning,
        ## Redirect message to the Debug stream
        [Parameter(ParameterSetName = 'Debug')]
        [System.Management.Automation.SwitchParameter] $IsDebug,
        ## Padding/indent section level
        [Parameter(ValueFromPipeline, ParameterSetName = 'Verbose')]
        [Parameter(ValueFromPipeline, ParameterSetName = 'Warning')]
        [Parameter(ValueFromPipeline, ParameterSetName = 'Debug')]
        [ValidateNotNullOrEmpty()]
        [System.Int16] $Indent
    )
    process {
        if ([System.String]::IsNullOrEmpty($Plugin)) {
            ## Attempt to resolve the plugin name from the parent scope
            if (Test-Path -Path Variable:\pluginName) { $Plugin = Get-Variable -Name pluginName -ValueOnly; }
            else { $Plugin = 'Unknown'; }
        }
        ## Center plugin name
        $pluginPaddingSize = [System.Math]::Floor((10 - $Plugin.Length) / 2);
        $pluginPaddingString = ''.PadRight($pluginPaddingSize);
        $Plugin = '{0}{1}' -f $pluginPaddingString, $Plugin;
        $Plugin = $Plugin.PadRight(10)
        $date = Get-Date;
        $sectionLevelPadding = ''.PadRight($Indent);
        $formattedMessage = '[ {0} ] [{1}] - {2}{3}' -f $date.ToString('HH:mm:ss:fff'), $Plugin, $sectionLevelPadding, $Message;
        switch ($PSCmdlet.ParameterSetName) {
            'Warning' { Write-Warning -Message $formattedMessage; }
            'Debug' { Write-Debug -Message $formattedMessage; }
            Default { Write-Verbose -Message $formattedMessage; }
        }
    } #end process
} #end function WriteLog

function Merge-PScriboPluginOption {
<#
    .SYNOPSIS
        Merges the specified options along with the plugin-specific default options.
#>
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param (
        ## Default/document options
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $DocumentOptions,
        ## Default plugin options to merge
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Collections.Hashtable] $DefaultPluginOptions,
        ## Specified runtime plugin options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $PluginOptions
    )
    process {
        $mergedOptions = $DocumentOptions.Clone();
        if ($null -ne $DefaultPluginOptions) {
            ## Overwrite the default document option with the plugin default option/value
            foreach ($option in $DefaultPluginOptions.GetEnumerator()) {
                $mergedOptions[$($option.Key)] = $option.Value;
            }
        }
        if ($null -ne $PluginOptions) {
            ## Overwrite the default document/plugin default option/value with the specified/runtime option
            foreach ($option in $PluginOptions.GetEnumerator()) {
                $mergedOptions[$($option.Key)] = $option.Value;
            }
        }
        return $mergedOptions;
    } #end process
} #end function

Function Test-CharsInPath {
<#
    .SYNOPSIS
    PowerShell function intended to verify if in the string what is the path to file or folder are incorrect chars.
    .DESCRIPTION
    PowerShell function intended to verify if in the string what is the path to file or folder are incorrect chars.
    Exit codes
    - 0 - everything OK
    - 1 - nothing to check
    - 2 - an incorrect char found in the path part
    - 3 - an incorrect char found in the file name part
    - 4 - incorrect chars found in the path part and in the file name part
    .PARAMETER Path
    Specifies the path to an item for what path (location on the disk) need to be checked.
    The Path can be an existing file or a folder on a disk provided as a PowerShell object or a string e.g. prepared to be used in file/folder creation.
    .PARAMETER SkipCheckCharsInFolderPart
    Skip checking in the folder part of path.
    .PARAMETER SkipCheckCharsInFileNamePart
    Skip checking in the file name part of path.
    .PARAMETER SkipDividingForParts
    Skip dividing provided path to a directory and a file name.
    Used usually in conjuction with SkipCheckCharsInFolderPart or SkipCheckCharsInFileNamePart.
    .EXAMPLE
    [PS] > Test-CharsInPath -Path $(Get-Item C:\Windows\Temp\new.csv') -Verbose
    VERBOSE: The path provided as a string was devided to, directory part: C:\Windows\Temp ; file name part: new.csv
    0
    Testing existing file. Returned code means that all chars are acceptable in the name of folder and file.
    .EXAMPLE
    [PS] > Test-CharsInPath -Path "C:\newfolder:2\nowy|.csv" -Verbose
    VERBOSE: The path provided as a string was devided to, directory part: C:\newfolder:2\ ; file name part: nowy|.csv
    VERBOSE: The incorrect char | with the UTF code [124] found in FileName part
    3
    Testing the string if can be used as a file name. The returned value means that can't do to an unsupported char in the file name.
    .OUTPUTS
    Exit code as an integer number. See description section to find the exit codes descriptions.
    .LINK
    https://github.com/it-praktyk/New-OutputObject
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, FileSystem
    REMARKS:
    # For Windows - based on the Power Tips
    # Finding Invalid File and Path Characters
    # http://community.idera.com/powershell/powertips/b/tips/posts/finding-invalid-file-and-path-characters
    # For PowerShell Core
    # https://docs.microsoft.com/en-us/dotnet/api/system.io.path.getinvalidpathchars?view=netcore-2.0
    # https://www.dwheeler.com/essays/fixing-unix-linux-filenames.html
    # [char]0 = NULL
    CURRENT VERSION
    - 0.6.1 - 2017-07-23
    HISTORY OF VERSIONS
    https://github.com/it-praktyk/New-OutputObject/CHANGELOG.md
#>
    [cmdletbinding()]
    [OutputType([System.Int32])]
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Path,
        [parameter(Mandatory = $false)]
        [switch]$SkipCheckCharsInFolderPart,
        [parameter(Mandatory = $false)]
        [switch]$SkipCheckCharsInFileNamePart,
        [parameter(Mandatory = $false)]
        [switch]$SkipDividingForParts
    )
    BEGIN {
        If (($PSVersionTable.ContainsKey('PSEdition')) -and ($PSVersionTable.PSEdition -eq 'Core') -and $IsLinux) {
            #[char]0 = NULL
            $PathInvalidChars = [char]0
            $FileNameInvalidChars = @([char]0, '/')
            $PathSeparators = @('/')
        }
        Elseif (($PSVersionTable.ContainsKey('PSEdition')) -and ($PSVersionTable.PSEdition -eq 'Core') -and $IsMacOS) {
            $PathInvalidChars = [char]58
            $FileNameInvalidChars = [char]58
            $PathSeparators = @('/')
        }
        #Windows
        Else {
            $PathInvalidChars = [System.IO.Path]::GetInvalidPathChars() #36 chars
            $FileNameInvalidChars = [System.IO.Path]::GetInvalidFileNameChars() #41 chars
            #$FileOnlyInvalidChars = @(':', '*', '?', '\', '/') #5 chars - as a difference
            $PathSeparators = @('/','\')
        }
        $IncorectCharFundInPath = $false
        $IncorectCharFundInFileName = $false
        $NothingToCheck = $true
    }
    END {
        [String]$DirectoryPath = ""
        [String]$FileName = ""
        $PathType = ($Path.GetType()).Name
        If (@('DirectoryInfo', 'FileInfo') -contains $PathType) {
            If (($SkipCheckCharsInFolderPart.IsPresent -and $PathType -eq 'DirectoryInfo') -or ($SkipCheckCharsInFileNamePart.IsPresent -and $PathType -eq 'FileInfo')) {
                Return 1
            }
            ElseIf ($PathType -eq 'DirectoryInfo') {
                [String]$DirectoryPath = $Path.FullName
            }
            elseif ($PathType -eq 'FileInfo') {
                [String]$DirectoryPath = $Path.DirectoryName
                [String]$FileName = $Path.Name
            }
        }
        ElseIf ($PathType -eq 'String') {
            If ( $SkipDividingForParts.IsPresent -and $SkipCheckCharsInFolderPart.IsPresent ) {
                $FileName = $Path
            }
            ElseIf ( $SkipDividingForParts.IsPresent -and $SkipCheckCharsInFileNamePart.IsPresent  ) {
                $DirectoryPath = $Path
            }
            Else {
                #Convert String to Array of chars
                $PathArray = $Path.ToCharArray()
                $PathLength = $PathArray.Length
                For ($i = ($PathLength-1); $i -ge 0; $i--) {
                    If ($PathSeparators -contains $PathArray[$i]) {
                        [String]$DirectoryPath = [String]$Path.Substring(0, $i +1)
                        break
                    }
                }
                If ([String]::IsNullOrEmpty($DirectoryPath)) {
                    [String]$FileName = [String]$Path
                }
                Else {
                    [String]$FileName = $Path.Replace($DirectoryPath, "")
                }
            }
        }
        Else {
            [String]$MessageText = "Input object {0} can't be tested" -f ($Path.GetType()).Name
            Throw $MessageText
        }
        [String]$MessageText = "The path provided as a string was divided to: directory part: {0} ; file name part: {1} ." -f $DirectoryPath, $FileName
        Write-Verbose -Message $MessageText
        If ($SkipCheckCharsInFolderPart.IsPresent -and $SkipCheckCharsInFileNamePart.IsPresent) {
            Return 1
        }
        If (-not ($SkipCheckCharsInFolderPart.IsPresent) -and -not [String]::IsNullOrEmpty($DirectoryPath)) {
            $NothingToCheck = $false
            foreach ($Char in $PathInvalidChars) {
                If ($DirectoryPath.ToCharArray() -contains $Char) {
                    $IncorectCharFundInPath = $true
                    [String]$MessageText = "The incorrect char {0} with the UTF code [{1}] found in the Path part." -f $Char, $([int][char]$Char)
                    Write-Verbose -Message $MessageText
                }
            }
        }
        If (-not ($SkipCheckCharsInFileNamePart.IsPresent) -and -not [String]::IsNullOrEmpty($FileName)) {
            $NothingToCheck = $false
            foreach ($Char in $FileNameInvalidChars) {
                If ($FileName.ToCharArray() -contains $Char) {
                    $IncorectCharFundInFileName = $true
                    [String]$MessageText = "The incorrect char {0} with the UTF code [{1}] found in FileName part." -f $Char, $([int][char]$Char)
                    Write-Verbose -Message $MessageText
                }
            }
        }
        If ($IncorectCharFundInPath -and $IncorectCharFundInFileName) {
            Return 4
        }
        elseif ($NothingToCheck) {
            Return 1
        }
        elseif ($IncorectCharFundInPath) {
            Return 2
        }
        elseif ($IncorectCharFundInFileName) {
            Return 3
        }
        Else {
            Return 0
        }
    }
}

function OutHtml {
<#
    .SYNOPSIS
        Html output plugin for PScribo.
    .DESCRIPTION
        Outputs a Html file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.IO.FileInfo])]
    param (
        ## PScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject] $Document,
        ## Output directory path for the .html file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path,
        ### Hashtable of all plugin supported options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {
        $pluginName = 'Html';
        #region OutHtml Private Functions
        function New-PScriboHtmlOption {
        <#
            .SYNOPSIS
                Sets the text plugin specific formatting/output options.
            .NOTES
                All plugin options should be prefixed with the plugin name.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Collections.Hashtable])]
            param (
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Boolean] $NoPageLayoutStyle = $false
            )
            process {
                return @{
                    NoPageLayoutStyle = $NoPageLayoutStyle;
                }
            } #end process
        } #end function New-PScriboHtmlOption
        function GetHtmlStyle {
        <#
            .SYNOPSIS
                Generates html stylesheet style attributes from a PScribo document style.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)] [System.Object] $Style
            )
            process {
                $styleBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $styleBuilder.AppendFormat(" font-family: '{0}';", $Style.Font -join "','");
                ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
                $invariantFontSize = ConvertToInvariantCultureString -Object ($Style.Size / 12) -Format 'f2';
                [ref] $null = $styleBuilder.AppendFormat(' font-size: {0}em;', $invariantFontSize);
                [ref] $null = $styleBuilder.AppendFormat(' text-align: {0};', $Style.Align.ToLower());
                if ($Style.Bold) {
                    [ref] $null = $styleBuilder.Append(' font-weight: bold;');
                }
                else {
                    [ref] $null = $styleBuilder.Append(' font-weight: normal;');
                }
                if ($Style.Italic) {
                    [ref] $null = $styleBuilder.Append(' font-style: italic;');
                }
                if ($Style.Underline) {
                    [ref] $null = $styleBuilder.Append(' text-decoration: underline;');
                }
                if ($Style.Color.StartsWith('#')) {
                    [ref] $null = $styleBuilder.AppendFormat(' color: {0};', $Style.Color.ToLower());
                }
                else {
                    [ref] $null = $styleBuilder.AppendFormat(' color: #{0};', $Style.Color);
                }
                if ($Style.BackgroundColor) {
                    if ($Style.BackgroundColor.StartsWith('#')) {
                        [ref] $null = $styleBuilder.AppendFormat(' background-color: {0};', $Style.BackgroundColor.ToLower());
                    }
                    else {
                        [ref] $null = $styleBuilder.AppendFormat(' background-color: #{0};', $Style.BackgroundColor.ToLower());
                    }
                }
                return $styleBuilder.ToString();
            }
        } #end function GetHtmlStyle
        function GetHtmlTableStyle {
        <#
            .SYNOPSIS
                Generates html stylesheet style attributes from a PScribo document table style.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document table style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TableStyle
            )
            process {
                $tableStyleBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $tableStyleBuilder.AppendFormat(' padding: {0}em {1}em {2}em {3}em;',
                    (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingTop)),
                        (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingRight)),
                            (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingBottom)),
                                (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingLeft))),
                [ref] $null = $tableStyleBuilder.AppendFormat(' border-style: {0};', $TableStyle.BorderStyle.ToLower());
                if ($TableStyle.BorderWidth -gt 0) {
                    $invariantBorderWidth = ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.BorderWidth);
                    [ref] $null = $tableStyleBuilder.AppendFormat(' border-width: {0}em;', $invariantBorderWidth);
                    if ($TableStyle.BorderColor.Contains('#')) {
                        [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: {0};', $TableStyle.BorderColor);
                    }
                    else {
                        [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: #{0};', $TableStyle.BorderColor);
                    }
                }
                [ref] $null = $tableStyleBuilder.Append(' border-collapse: collapse;');
                ## <table align="center"> is deprecated in Html5
                if ($TableStyle.Align -eq 'Center') {
                    [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: auto;');
                }
                elseif ($TableStyle.Align -eq 'Right') {
                    [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: 0;');
                }
                return $tableStyleBuilder.ToString();
            }
        } #end function Outhtmltablestyle
        function GetHtmlTableDiv {
        <#
            .SYNOPSIS
                Generates Html <div style=..><table style=..> tags based upon table width, columns and indentation
            .NOTES
                A <div> is required to ensure that the table stays within the "page" boundaries/margins.
        #>
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {
                $divBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.Tabs -gt 0) {
                    $invariantMarginLeft = ConvertToInvariantCultureString -Object (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs));
                    [ref] $null = $divBuilder.AppendFormat('<div style="margin-left: {0}em;">' -f $invariantMarginLeft);
                }
                else {
                    [ref] $null = $divBuilder.Append('<div>' -f (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)));
                }
                if ($Table.List) {
                    [ref] $null = $divBuilder.AppendFormat('<table class="{0}-list"', $Table.Style.ToLower());
                }
                else {
                    [ref] $null = $divBuilder.AppendFormat('<table class="{0}"', $Table.Style.ToLower());
                }
                $styleElements = @();
                if ($Table.Width -gt 0) {
                    $styleElements += 'width:{0}%;' -f $Table.Width;
                }
                if ($Table.ColumnWidths) {
                    $styleElements += 'table-layout: fixed;';
                    $styleElements += 'word-break: break-word;'
                }
                if ($styleElements.Count -gt 0) {
                    [ref] $null = $divBuilder.AppendFormat(' style="{0}">', [String]::Join(' ', $styleElements));
                }
                else {
                    [ref] $null = $divBuilder.Append('>');
                }
                return $divBuilder.ToString();
            }
        } #end function GetHtmlTableDiv
        function GetHtmlTableColGroup {
        <#
            .SYNOPSIS
                Generates Html <colgroup> tags based on table column widths
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {
                $colGroupBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.ColumnWidths) {
                    [ref] $null = $colGroupBuilder.Append('<colgroup>');
                    foreach ($columnWidth in $Table.ColumnWidths) {
                        if ($null -eq $columnWidth) {
                            [ref] $null = $colGroupBuilder.Append('<col />');
                        }
                        else {
                            [ref] $null = $colGroupBuilder.AppendFormat('<col style="max-width:{0}%; min-width:{0}%; width:{0}%" />', $columnWidth);
                        }
                    }
                    [ref] $null = $colGroupBuilder.AppendLine('</colgroup>');
                }
                return $colGroupBuilder.ToString();
            }
        } #end function GetHtmlTableDiv
        function OutHtmlTOC {
        <#
            .SYNOPSIS
                Generates Html table of contents.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC
            )
            process {
                $tocBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $tocBuilder.AppendFormat('<h1 class="{0}">{1}</h1>', $TOC.ClassId, $TOC.Name);
                #[ref] $null = $tocBuilder.AppendLine('<table style="width: 100%;">');
                [ref] $null = $tocBuilder.AppendLine('<table>');
                foreach ($tocEntry in $Document.TOC) {
                    $sectionNumberIndent = '&nbsp;&nbsp;&nbsp;' * $tocEntry.Level;
                    if ($Document.Options['EnableSectionNumbering']) {
                        [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}</td><td>{1}<a href="#{2}" style="text-decoration: none;">{3}</a></td></tr>', $tocEntry.Number, $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine();
                    }
                    else {
                        [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}<a href="#{1}" style="text-decoration: none;">{2}</a></td></tr>', $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine();
                    }
                }
                [ref] $null = $tocBuilder.AppendLine('</table>');
                return $tocBuilder.ToString();
            } #end process
        } #end function OutHtmlTOC
        function OutHtmlBlankLine {
        <#
            .SYNOPSIS
                Outputs html PScribo.Blankline.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $BlankLine
            )
            process {
                $blankLineBuilder = New-Object -TypeName System.Text.StringBuilder;
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
                    [ref] $null = $blankLineBuilder.Append('<br />');
                }
                return $blankLineBuilder.ToString();
            } #end process
        } #end function OutHtmlBlankLine
        function OutHtmlStyle {
        <#
            .SYNOPSIS
                Generates an in-line HTML CSS stylesheet from a PScribo document styles and table styles.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document styles
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [System.Collections.Hashtable] $Styles,
                ## PScribo document tables styles
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [System.Collections.Hashtable] $TableStyles,
                ## Suppress page layout styling
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $NoPageLayoutStyle
            )
            process {
                $stylesBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $stylesBuilder.AppendLine('<style type="text/css">');
                if (-not $NoPageLayoutStyle) {
                    ## Add HTML page layout styling options, e.g. when emailing HTML documents
                    [ref] $null = $stylesBuilder.AppendLine('html { height: 100%; -webkit-background-size: cover; -moz-background-size: cover; -o-background-size: cover; background-size: cover; background: #f8f8f8; }');
                    [ref] $null = $stylesBuilder.Append("page { background: white; width: $($Document.Options['PageWidth'])mm; display: block; margin-top: 1em; margin-left: auto; margin-right: auto; margin-bottom: 1em; ");
                    [ref] $null = $stylesBuilder.AppendLine('border-style: solid; border-width: 1px; border-color: #c6c6c6; }');
                    [ref] $null = $stylesBuilder.AppendLine('@media print { body, page { margin: 0; box-shadow: 0; } }');
                    [ref] $null = $stylesBuilder.AppendLine('hr { margin-top: 1.0em; }');
                }
                foreach ($style in $Styles.Keys) {
                    ## Build style
                    $htmlStyle = GetHtmlStyle -Style $Styles[$style];
                    [ref] $null = $stylesBuilder.AppendFormat(' .{0} {{{1} }}', $Styles[$style].Id, $htmlStyle).AppendLine();
                }
                foreach ($tableStyle in $TableStyles.Keys) {
                    $tStyle = $TableStyles[$tableStyle];
                    $tableStyleId = $tStyle.Id.ToLower();
                    $htmlTableStyle = GetHtmlTableStyle -TableStyle $tStyle;
                    $htmlHeaderStyle = GetHtmlStyle -Style $Styles[$tStyle.HeaderStyle];
                    $htmlRowStyle = GetHtmlStyle -Style $Styles[$tStyle.RowStyle];
                    $htmlAlternateRowStyle = GetHtmlStyle -Style $Styles[$tStyle.AlternateRowStyle];
                    ## Generate Standard table styles
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} th {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(odd) td {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(even) td {{{1}{2} }}', $tableStyleId, $htmlAlternateRowStyle, $htmlTableStyle).AppendLine();
                    ## Generate List table styles
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list td:nth-child(1) {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list td:nth-child(2) {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine();
                } #end foreach style
                [ref] $null = $stylesBuilder.AppendLine('</style>');
                return $stylesBuilder.ToString().TrimEnd();
            } #end process
        } #end function OutHtmlStyle
        function OutHtmlSection {
        <#
            .SYNOPSIS
                Output formatted Html section.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## Section to output
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            process {
                [System.Text.StringBuilder] $sectionBuilder = New-Object System.Text.StringBuilder;
                $encodedSectionName = [System.Net.WebUtility]::HtmlEncode($Section.Name);
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $encodedSectionName; }
                else { [string] $sectionName = '{0}' -f $encodedSectionName; }
                [int] $headerLevel = $Section.Number.Split('.').Count;
                ## Html <h5> is the maximum supported level
                if ($headerLevel -gt 6) {
                    WriteLog -Message $localized.MaxHeadingLevelWarning -IsWarning;
                    $headerLevel = 6;
                }
                if ([string]::IsNullOrEmpty($Section.Style)) {
                    $className = $Document.DefaultStyle;
                }
                else {
                    $className = $Section.Style;
                }
                [ref] $null = $sectionBuilder.AppendFormat('<a name="{0}"><h{1} class="{2}">{3}</h{1}></a>', $Section.Id, $headerLevel, $className, $sectionName.TrimStart());
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) {
                        $sectionId = '{0}[..]' -f $s.Id.Substring(0,36);
                    }
                    else {
                        $sectionId = $s.Id;
                    }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) {
                        $currentIndentationLevel = $s.Level +1;
                    }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutHtmlSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append((OutHtmlParagraph -Paragraph $s)); }
                        'PScribo.LineBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlLineBreak)); }
                        'PScribo.PageBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlPageBreak)); }
                        'PScribo.Table' { [ref] $null = $sectionBuilder.Append((OutHtmlTable -Table $s)); }
                        'PScribo.BlankLine' { [ref] $null = $sectionBuilder.Append((OutHtmlBlankLine -BlankLine $s)); }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
                return $sectionBuilder.ToString();
            } #end process
        } # end function OutHtmlSection
        function GetHtmlParagraphStyle {
        <#
            .SYNOPSIS
                Generates html style attribute from PScribo paragraph style overrides.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()] [System.Object] $Paragraph
            )
            process {
                $paragraphStyleBuilder = New-Object -TypeName System.Text.StringBuilder;
                if ($Paragraph.Tabs -gt 0) {
                    ## Default to 1/2in tab spacing
                    $tabEm = ConvertToInvariantCultureString -Object (ConvertMmToEm -Millimeter (12.7 * $Paragraph.Tabs)) -Format 'f2';
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' margin-left: {0}em;', $tabEm);
                }
                if ($Paragraph.Font) {
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(" font-family: '{0}';", $Paragraph.Font -Join "','");
                }
                if ($Paragraph.Size -gt 0) {
                    ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
                    $invariantParagraphSize = ConvertToInvariantCultureString -Object ($Paragraph.Size / 12) -Format 'f2';
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' font-size: {0}em;', $invariantParagraphSize);
                }
                if ($Paragraph.Bold -eq $true) {
                    [ref] $null = $paragraphStyleBuilder.Append(' font-weight: bold;');
                }
                if ($Paragraph.Italic -eq $true) {
                    [ref] $null = $paragraphStyleBuilder.Append(' font-style: italic;');
                }
                if ($Paragraph.Underline -eq $true) {
                    [ref] $null = $paragraphStyleBuilder.Append(' text-decoration: underline;');
                }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Color) -and $Paragraph.Color.StartsWith('#')) {
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: {0};', $Paragraph.Color.ToLower());
                }
                elseif (-not [System.String]::IsNullOrEmpty($Paragraph.Color)) {
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: #{0};', $Paragraph.Color.ToLower());
                }
                return $paragraphStyleBuilder.ToString().TrimStart();
            } #end process
        } #end function GetHtmlParagraphStyle
        function OutHtmlParagraph {
        <#
            .SYNOPSIS
                Output formatted Html paragraph.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            process {
                [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                $encodedText = [System.Net.WebUtility]::HtmlEncode($Paragraph.Text);
                if ([System.String]::IsNullOrEmpty($encodedText)) {
                    $encodedText = [System.Net.WebUtility]::HtmlEncode($Paragraph.Id);
                }
                # $encodedText = $encodedText -replace [System.Environment]::NewLine, '<br />';
                $encodedText = $encodedText.Replace([System.Environment]::NewLine, '<br />');
                $customStyle = GetHtmlParagraphStyle -Paragraph $Paragraph;
                if ([System.String]::IsNullOrEmpty($Paragraph.Style) -and [System.String]::IsNullOrEmpty($customStyle)) {
                    [ref] $null = $paragraphBuilder.AppendFormat('<div>{0}</div>', $encodedText);
                }
                elseif ([System.String]::IsNullOrEmpty($customStyle)) {
                    [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}">{1}</div>', $Paragraph.Style, $encodedText);
                }
                else {
                    [ref] $null = $paragraphBuilder.AppendFormat('<div style="{1}">{2}</div>', $Paragraph.Style, $customStyle, $encodedText);
                }
                return $paragraphBuilder.ToString();
            } #end process
        } #end OutHtmlParagraph
        function GetHtmlTableList {
        <#
            .SYNOPSIS
                Generates list html <table> from a PScribo.Table row object.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,
                [Parameter(Mandatory)]
                [System.Object] $Row
            )
            process {
                $listTableBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $listTableBuilder.Append((GetHtmlTableDiv -Table $Table));
                [ref] $null = $listTableBuilder.Append((GetHtmlTableColGroup -Table $Table));
                [ref] $null = $listTableBuilder.Append('<tbody>');
                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                    $propertyName = $Table.Columns[$i];
                    [ref] $null = $listTableBuilder.AppendFormat('<tr><td>{0}</td>', $propertyName);
                    $propertyStyle = '{0}__Style' -f $propertyName;
                    if ($row.PSObject.Properties[$propertyStyle]) {
                        $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle]);
                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {
                            [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">&nbsp;</td></tr>', $propertyStyleHtml);
                        }
                        else {
                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString());
                            $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />');
                            [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">{1}</td></tr>', $propertyStyleHtml, $encodedHtmlContent);
                        }
                    }
                    else {
                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {
                            [ref] $null = $listTableBuilder.Append('<td>&nbsp;</td></tr>');
                        }
                        else {
                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString());
                            $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />')
                            [ref] $null = $listTableBuilder.AppendFormat('<td>{0}</td></tr>', $encodedHtmlContent);
                        }
                    }
                } #end for each property
                [ref] $null = $listTableBuilder.AppendLine('</tbody></table></div>');
                return $listTableBuilder.ToString();
            } #end process
        } #end function GetHtmlTableList
        function GetHtmlTable {
        <#
            .SYNOPSIS
                Generates html <table> from a PScribo.Table object.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {
                $standardTableBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $standardTableBuilder.Append((GetHtmlTableDiv -Table $Table));
                [ref] $null = $standardTableBuilder.Append((GetHtmlTableColGroup -Table $Table));
                ## Table headers
                [ref] $null = $standardTableBuilder.Append('<thead><tr>');
                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                    [ref] $null = $standardTableBuilder.AppendFormat('<th>{0}</th>', $Table.Columns[$i]);
                }
                [ref] $null = $standardTableBuilder.Append('</tr></thead>');
                ## Table body
                [ref] $null = $standardTableBuilder.AppendLine('<tbody>');
                foreach ($row in $Table.Rows) {
                    [ref] $null = $standardTableBuilder.Append('<tr>');
                    foreach ($propertyName in $Table.Columns) {
                        $propertyStyle = '{0}__Style' -f $propertyName;
                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {
                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode('&nbsp;');
                        }
                        else {
                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString());
                        }
                        $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />');
                        if ($row.PSObject.Properties[$propertyStyle]) {
                            ## Cell styles override row styles
                            $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]).Trim();
                            [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $encodedHtmlContent);
                        }
                        elseif (($row.PSObject.Properties['__Style']) -and (-not [System.String]::IsNullOrEmpty($row.__Style))) {
                            ## We have a row style
                            $rowStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.__Style]).Trim();
                            [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $rowStyleHtml, $encodedHtmlContent);
                        }
                        else {
                            if ($null -ne $row.$propertyName) {
                                ## Check that the property has a value
                                [ref] $null = $standardTableBuilder.AppendFormat('<td>{0}</td>', $encodedHtmlContent);
                            }
                            else {
                                [ref] $null = $standardTableBuilder.Append('<td>&nbsp</td>');
                            }
                        } #end if $row.PropertyStyle
                    } #end foreach property
                    [ref] $null = $standardTableBuilder.AppendLine('</tr>');
                } #end foreach row
                [ref] $null = $standardTableBuilder.AppendLine('</tbody></table></div>');
                return $standardTableBuilder.ToString();
            } #end process
        } #end function GetHtmlTableList
        function OutHtmlTable {
        <#
            .SYNOPSIS
                Output formatted Html <table> from PScribo.Table object.
            .NOTES
                One table is output per table row with the -List parameter.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()] [System.Object] $Table
            )
            process {
                [System.Text.StringBuilder] $tableBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.List) {
                    ## Create a table for each row
                    for ($r = 0; $r -lt $Table.Rows.Count; $r++) {
                        $row = $Table.Rows[$r];
                        if ($r -gt 0) {
                            ## Add a space between each table to mirror Word output rendering
                            [ref] $null = $tableBuilder.AppendLine('<p />');
                        }
                        [ref] $null = $tableBuilder.Append((GetHtmlTableList -Table $Table -Row $row));
                    } #end foreach row
                }
                else {
                    [ref] $null = $tableBuilder.Append((GetHtmlTable -Table $Table));
                } #end if
                return $tableBuilder.ToString();
                #Write-Output ($tableBuilder.ToString()) -NoEnumerate;
            } #end process
        } #end function outhtmltable
        function OutHtmlLineBreak {
        <#
            .SYNOPSIS
                Output formatted Html line break.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {
                return '<hr />';
            }
        } #end function OutHtmlLineBreak
        function OutHtmlPageBreak {
        <#
            .SYNOPSIS
                Output formatted Html page break.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {
                [System.Text.StringBuilder] $pageBreakBuilder = New-Object 'System.Text.StringBuilder';
                [ref] $null = $pageBreakBuilder.Append('</div></page>');
                $topMargin = ConvertMmToEm $Document.Options['MarginTop'];
                $leftMargin = ConvertMmToEm $Document.Options['MarginLeft'];
                $bottomMargin = ConvertMmToEm $Document.Options['MarginBottom'];
                $rightMargin = ConvertMmToEm $Document.Options['MarginRight'];
                [ref] $null = $pageBreakBuilder.AppendFormat('<page><div class="{0}" style="padding-top: {1}em; padding-left: {2}em; padding-bottom: {3}em; padding-right: {4}em;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine();
                return $pageBreakBuilder.ToString();
            }
        } #end function OutHtmlPageBreak
        #endregion OutHtml Private Functions
    } #end begin
    process {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);
        ## Merge the document, plugin default and specified/specific plugin options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboHtmlOption;
            DocumentOptions = $Document.Options;
            PluginOptions = $Options;
        }
        $options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams;
        $noPageLayoutStyle = $Options['NoPageLayoutStyle'];
        $topMargin = ConvertMmToEm -Millimeter $options['MarginTop'];
        $leftMargin = ConvertMmToEm -Millimeter $options['MarginLeft'];
        $bottomMargin = ConvertMmToEm -Millimeter $options['MarginBottom'];
        $rightMargin = ConvertMmToEm -Millimeter $options['MarginRight'];
        [System.Text.StringBuilder] $htmlBuilder = New-Object System.Text.StringBuilder;
        [ref] $null = $htmlBuilder.AppendLine('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">');
        [ref] $null = $htmlBuilder.AppendLine('<html xmlns="http://www.w3.org/1999/xhtml">');
        [ref] $null = $htmlBuilder.AppendLine('<head><title>{0}</title>' -f $Document.Name);
        [ref] $null = $htmlBuilder.AppendLine('{0}</head><body><page>' -f (OutHtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles -NoPageLayoutStyle:$noPageLayoutStyle));
        [ref] $null = $htmlBuilder.AppendFormat('<div class="{0}" style="padding-top: {1}em; padding-left: {2}em; padding-bottom: {3}em; padding-right: {4}em;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine();
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { [ref] $null = $htmlBuilder.Append((OutHtmlSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $htmlBuilder.Append((OutHtmlParagraph -Paragraph $s)); }
                'PScribo.Table' { [ref] $null = $htmlBuilder.Append((OutHtmlTable -Table $s)); }
                'PScribo.LineBreak' { [ref] $null = $htmlBuilder.Append((OutHtmlLineBreak)); }
                'PScribo.PageBreak' { [ref] $null = $htmlBuilder.Append((OutHtmlPageBreak)); } ## Page breaks are implemented as line breaks with extra padding
                'PScribo.TOC' { [ref] $null = $htmlBuilder.Append((OutHtmlTOC -TOC $s)); }
                'PScribo.BlankLine' { [ref] $null = $htmlBuilder.Append((OutHtmlBlankLine -BlankLine $s)); }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            } #end switch
        } #end foreach section
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        $destinationPath = Join-Path $Path ('{0}.html' -f $Document.Name);
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        $htmlBuilder.ToString().TrimEnd() | Out-File -FilePath $destinationPath -Force -Encoding utf8;
        [ref] $null = $htmlBuilder;
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        Write-Output (Get-Item -Path $destinationPath);
    } #end process
} #end function OutHtml

function OutText {
<#
    .SYNOPSIS
        Text output plugin for PScribo.
    .DESCRIPTION
        Outputs a text file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    param (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,
        ## Output directory path for the .txt file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,
        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {
        $pluginName = 'Text';
        #region OutText Private Functions
        function New-PScriboTextOption {
        <#
            .SYNOPSIS
                Sets the text plugin specific formatting/output options.
            .NOTES
                All plugin options should be prefixed with the plugin name.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Collections.Hashtable])]
            param (
                ## Text/output width. 0 = none/no wrap.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Int32] $TextWidth = 120,
                ## Document header separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $HeaderSeparator = '=',
                ## Document section separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $SectionSeparator = '-',
                ## Document section separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $LineBreakSeparator = '_',
                ## Default header/section separator width.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Int32] $SeparatorWidth = $TextWidth,
                ## Text encoding
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('ASCII','Unicode','UTF7','UTF8')]
                [System.String] $Encoding = 'ASCII'
            )
            process {
                return @{
                    TextWidth = $TextWidth;
                    HeaderSeparator = $HeaderSeparator;
                    SectionSeparator = $SectionSeparator;
                    LineBreakSeparator = $LineBreakSeparator;
                    SeparatorWidth = $SeparatorWidth;
                    Encoding = $Encoding;
                }
            } #end process
        } #end function New-PScriboTextOption
        function OutTextTOC {
        <#
            .SYNOPSIS
                Output formatted Table of Contents
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC
            )
            begin {
                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {
                    $options = New-PScriboTextOption;
                }
            }
            process {
                $tocBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $tocBuilder.AppendLine($TOC.Name);
                [ref] $null = $tocBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator));
                if ($Options.ContainsKey('EnableSectionNumbering')) {
                    $maxSectionNumberLength = ([System.String] ($Document.TOC.Number | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum)).Length;
                    foreach ($tocEntry in $Document.TOC) {
                        $sectionNumberPaddingLength = $maxSectionNumberLength - $tocEntry.Number.Length;
                        $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ');
                        $sectionPadding = ''.PadRight($sectionNumberPaddingLength, ' ');
                        [ref] $null = $tocBuilder.AppendFormat('{0}{1}  {2}{3}', $tocEntry.Number, $sectionPadding, $sectionNumberIndent, $tocEntry.Name).AppendLine();
                    } #end foreach TOC entry
                }
                else {
                    $maxSectionNumberLength = $Document.TOC.Level | Sort-Object | Select-Object -Last 1;
                    foreach ($tocEntry in $Document.TOC) {
                        $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ');
                        [ref] $null = $tocBuilder.AppendFormat('{0}{1}', $sectionNumberIndent, $tocEntry.Name).AppendLine();
                    } #end foreach TOC entry
                }
                return $tocBuilder.ToString();
            } #end process
        } #end function OutTextTOC
        function OutTextBlankLine {
        <#
            .SYNOPSIS
                Output formatted text blankline.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $BlankLine
            )
            process {
                $blankLineBuilder = New-Object -TypeName System.Text.StringBuilder;
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
                    [ref] $null = $blankLineBuilder.AppendLine();
                }
                return $blankLineBuilder.ToString();
            } #end process
        } #end function OutHtmlBlankLine
        function OutTextSection {
        <#
            .SYNOPSIS
                Output formatted text section.
        #>
            [CmdletBinding()]
            param (
                ## Section to output
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            begin {
                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {
                    $options = New-PScriboTextOption;
                }
            }
            process {
                $sectionBuilder = New-Object -TypeName System.Text.StringBuilder;
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name; }
                else { [string] $sectionName = '{0}' -f $Section.Name; }
                [ref] $null = $sectionBuilder.AppendLine();
                [ref] $null = $sectionBuilder.AppendLine($sectionName.TrimStart());
                [ref] $null = $sectionBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator));
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutTextSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append(($s | OutTextParagraph)); }
                        'PScribo.PageBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextPageBreak)); }  ## Page breaks implemented as line break with extra padding
                        'PScribo.LineBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextLineBreak)); }
                        'PScribo.Table' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextTable)); }
                        'PScribo.BlankLine' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextBlankLine)); }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
                return $sectionBuilder.ToString();
            } #end process
        } #end function outtextsection
        function OutTextParagraph {
        <#
            .SYNOPSIS
                Output formatted paragraph text.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            begin {
                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {
                    $options = New-PScriboTextOption;
                }
            }
            process {
                $padding = ''.PadRight(($Paragraph.Tabs * 4), ' ');
                if ([string]::IsNullOrEmpty($Paragraph.Text)) { $text = "$padding$($Paragraph.Id)"; }
                else { $text = "$padding$($Paragraph.Text)"; }
                $formattedText = OutStringWrap -InputObject $text -Width $Options.TextWidth;
                if ($Paragraph.NewLine) { return "$formattedText`r`n"; }
                else { return $formattedText; }
            } #end process
        } #end outtextparagraph
        function OutTextLineBreak {
        <#
            .SYNOPSIS
                Output formatted line break text.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            begin {
                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {
                    $options = New-PScriboTextOption;
                }
            }
            process {
                ## Use the specified output width
                if ($options.TextWidth -eq 0) { $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1; }
                $lb = ''.PadRight($options.SeparatorWidth, $options.LineBreakSeparator);
                return "$(OutStringWrap -InputObject $lb -Width $options.TextWidth)`r`n";
            } #end process
        } #end function OutTextLineBreak
        function OutTextPageBreak {
        <#
            .SYNOPSIS
                Output formatted line break text.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {
                return "$(OutTextLineBreak)`r`n";
            } #end process
        } #end function OutTextLineBreak
        function OutTextTable {
        <#
            .SYNOPSIS
                Output formatted text table.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Table
            )
            begin {
                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {
                    $options = New-PScriboTextOption;
                }
            }
            process {
                ## Use the current output buffer width
                if ($options.TextWidth -eq 0) { $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1; }
                if ($Table.List) {
                    $text = ($Table.Rows | Select-Object -Property * -ExcludeProperty '*__Style' | Format-List | Out-String -Width $options.TextWidth).Trim();
                } else {
                    ## Don't trim tabs for table headers
                    ## Tables set to AutoSize as otherwise, rendering is different between PoSh v4 and v5
                    $text = ($Table.Rows | Select-Object -Property * -ExcludeProperty '*__Style' | Format-Table -Wrap -AutoSize | Out-String -Width $options.TextWidth).Trim("`r`n");
                }
                # Ensure there's a space before and after the table.
                return "`r`n$text`r`n";
            } #end process
        } #end function outtexttable
        function OutStringWrap {
        <#
            .SYNOPSIS
                Outputs objects to strings, wrapping as required.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [Object[]] $InputObject,
                [Parameter()]
                [ValidateNotNull()]
                [System.Int32] $Width = $Host.UI.RawUI.BufferSize.Width
            )
            begin {
                ## 2 is the minimum, therefore default to wiiiiiiiiiide!
                if ($Width -lt 2) { $Width = 4096; }
                WriteLog -Message ('Wrapping text at "{0}" characters.' -f $Width) -IsDebug;
            }
            process {
                foreach ($object in $InputObject) {
                    $textBuilder = New-Object -TypeName System.Text.StringBuilder;
                    $text = (Out-String -InputObject $object).TrimEnd("`r`n");
                    for ($i = 0; $i -le $text.Length; $i += $Width) {
                        if (($i + $Width) -ge ($text.Length -1)) { [ref] $null = $textBuilder.Append($text.Substring($i)); }
                        else { [ref] $null = $textBuilder.AppendLine($text.Substring($i, $Width)); }
                    } #end for
                    return $textBuilder.ToString();
                    $textBuilder = $null;
                } #end foreach
            } #end process
        } #end function OutStringWrap
        #endregion OutText Private Functions
    }
    process {
        $stopwatch = [Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);
        ## Merge the document, text default and specified text options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboTextOption;
            DocumentOptions = $Document.Options;
            PluginOptions = $Options;
        }
        $Options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams;
        [System.Text.StringBuilder] $textBuilder = New-Object System.Text.StringBuilder;
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { [ref] $null = $textBuilder.Append((OutTextSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $textBuilder.Append(($s | OutTextParagraph)); }
                'PScribo.PageBreak' { [ref] $null = $textBuilder.AppendLine((OutTextPageBreak)); }
                'PScribo.LineBreak' { [ref] $null = $textBuilder.AppendLine((OutTextLineBreak)); }
                'PScribo.Table' { [ref] $null = $textBuilder.AppendLine(($s | OutTextTable)); }
                'PScribo.TOC' { [ref] $null = $textBuilder.AppendLine(($s | OutTextTOC)); }
                'PScribo.BlankLine' { [ref] $null = $textBuilder.AppendLine(($s | OutTextBlankLine)); }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            } #end switch
        } #end foreach
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        $destinationPath = Join-Path -Path $Path ('{0}.txt' -f $Document.Name);
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        Set-Content -Value ($textBuilder.ToString()) -Path $destinationPath -Encoding $Options.Encoding;
        [ref] $null = $textBuilder;
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath);
    } #end process
} #end function OutText

function OutWord {
<#
    .SYNOPSIS
        Microsoft Word output plugin for PScribo.
    .DESCRIPTION
        Outputs a Word document representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.IO.FileInfo])]
    param (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,
        ## Output directory path for the .txt file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,
        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {
        $pluginName = 'Word';
        #region OutWord Private Functions
        function ConvertToWordColor {
        <#
            .SYNOPSIS
                Converts an HTML color to RRGGBB value as Word does not support short Html color codes
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.String] $Color
            )
            process {
                $Color = $Color.TrimStart('#');
                if ($Color.Length -eq 3) {
                    $Color = '{0}{0}{1}{1}{2}{2}' -f $Color[0], $Color[1],$Color[2];
                }
                return $Color.ToUpper();
            }
        } #end function ConvertToWordColor
        function OutWordSection {
        <#
            .SYNOPSIS
                Output formatted Word section (paragraph).
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section,
                [Parameter(Mandatory)]
                [System.Xml.XmlElement] $RootElement,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $p = $RootElement.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                if (-not [System.String]::IsNullOrEmpty($Section.Style)) {
                    #if (-not $Section.IsExcluded) {
                        ## If it's excluded we need a non-Heading style :( Could explicitly set the style on the run?
                        $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                        [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Section.Style);
                    #}
                }
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                ## Increment heading spacing by 2pt for each section level, starting at 8pt for level 0, 10pt for level 1 etc
                $spacingPt = (($Section.Level * 2) + 8) * 20;
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, $spacingPt);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, $spacingPt);
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                if ($Document.Options['EnableSectionNumbering']) {
                    [System.String] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name;
                }
                else {
                    [System.String] $sectionName = '{0}' -f $Section.Name;
                }
                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($sectionName));
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) {
                        $sectionId = '{0}[..]' -f $s.Id.Substring(0,36);
                    }
                    else {
                        $sectionId = $s.Id;
                    }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) {
                        $currentIndentationLevel = $s.Level +1;
                    }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { $s | OutWordSection -RootElement $RootElement -XmlDocument $XmlDocument; }
                        'PScribo.Paragraph' { [ref] $null = $RootElement.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $XmlDocument)); }
                        'PScribo.PageBreak' { [ref] $null = $RootElement.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $xmlDocument)); }
                        'PScribo.LineBreak' { [ref] $null = $RootElement.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $xmlDocument)); }
                        'PScribo.Table' { OutWordTable -Table $s -XmlDocument $xmlDocument -Element $RootElement; }
                        'PScribo.BlankLine' { OutWordBlankLine -BlankLine $s -XmlDocument $xmlDocument -Element $RootElement; }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
            } #end process
        } #end function OutWordSection
        function OutWordParagraph {
        <#
            .SYNOPSIS
                Output formatted Word paragraph.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Paragraph,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                if ($Paragraph.Tabs -gt 0) {
                    $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlnsMain));
                    [ref] $null = $ind.SetAttribute('left', $xmlnsMain, (720 * $Paragraph.Tabs));
                }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Style)) {
                    $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                    [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Paragraph.Style);
                }
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0);
                if ([System.String]::IsNullOrEmpty($Paragraph.Text)) {
                    $lines = $Paragraph.Id -Split [System.Environment]::NewLine;
                }
                else {
                    $lines = $Paragraph.TexT -Split [System.Environment]::NewLine;
                }
                ## Create a separate run for each line/break
                for ($l = 0; $l -lt $lines.Count; $l++) {
                    $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                    $rPr = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain));
                    ## Apply custom paragraph styles to the run..
                    if ($Paragraph.Font) {
                        $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain));
                        [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Paragraph.Font[0]);
                        [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Paragraph.Font[0]);
                    }
                    if ($Paragraph.Size -gt 0) {
                        $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain));
                        [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Paragraph.Size * 2);
                    }
                    if ($Paragraph.Bold -eq $true) {
                        [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain));
                    }
                    if ($Paragraph.Italic -eq $true) {
                        [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain));
                    }
                    if ($Paragraph.Underline -eq $true) {
                        $u = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain));
                        [ref] $null = $u.SetAttribute('val', $xmlnsMain, 'single');
                    }
                    if (-not [System.String]::IsNullOrEmpty($Paragraph.Color)) {
                        $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain));
                        [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Paragraph.Color));
                    }
                    $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                    [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve'); ## needs to be xml:space="preserve" NOT w:space...
                    [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($lines[$l]));
                    if ($l -lt ($lines.Count -1)) {
                        ## Don't add a line break to the last line/break
                        $brr = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                        $brt = $brr.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                        [ref] $null = $brt.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                    }
                } #end foreach line break
                return $p;
            } #end process
        } #end function OutWordParagraph
        function OutWordPageBreak {
            <#
            .SYNOPSIS
                Output formatted Word page break.
            #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $PageBreak,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $br = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                [ref] $null = $br.SetAttribute('type', $xmlnsMain, 'page');
                return $p;
            }
        } #end function OutWordPageBreak
        function OutWordLineBreak {
        <#
            .SYNOPSIS
                Output formatted Word line break.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $LineBreak,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $pBdr = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pBdr', $xmlnsMain));
                $bottom = $pBdr.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain));
                [ref] $null = $bottom.SetAttribute('val', $xmlnsMain, 'single');
                [ref] $null = $bottom.SetAttribute('sz', $xmlnsMain, 6);
                [ref] $null = $bottom.SetAttribute('space', $xmlnsMain, 1);
                [ref] $null = $bottom.SetAttribute('color', $xmlnsMain, 'auto');
                return $p;
            }
        } #end function OutWordLineBreak
        function GetWordTable {
        <#
            .SYNOPSIS
                Creates a scaffold Word <w:tbl> element
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tableStyle = $Document.TableStyles[$Table.Style];
                $tbl = $XmlDocument.CreateElement('w', 'tbl', $xmlnsMain);
                $tblPr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain));
                if ($Table.Tabs -gt 0) {
                    $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblInd', $xmlnsMain));
                    [ref] $null = $tblInd.SetAttribute('w', $xmlnsMain, (720 * $Table.Tabs));
                }
                if ($Table.ColumnWidths) {
                    $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain));
                    [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'fixed');
                }
                elseif ($Table.Width -eq 0) {
                    $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain));
                    [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'autofit');
                }
                if ($Table.Width -gt 0) {
                    $tblW = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblW', $xmlnsMain));
                    [ref] $null = $tblW.SetAttribute('type', $xmlnsMain, 'pct');
                    $tableWidthRenderPct = $Table.Width;
                    if ($Table.Tabs -gt 0) {
                        ## We now need to deal with tables being pushed outside the page margin
                        $pageWidthMm = $Document.Options['PageWidth'] - ($Document.Options['PageMarginLeft'] + $Document.Options['PageMarginRight']);
                        $indentWidthMm = ConvertPtToMm -Point ($Table.Tabs * 36);
                        $tableRenderMm = (($pageWidthMm / 100) * $Table.Width) + $indentWidthMm;
                        if ($tableRenderMm -gt $pageWidthMm) {
                            ## We've over-flowed so need to work out the maximum percentage
                            $maxTableWidthMm = $pageWidthMm - $indentWidthMm;
                            $tableWidthRenderPct = [System.Math]::Round(($maxTableWidthMm / $pageWidthMm) * 100, 2);
                            WriteLog -Message ($localized.TableWidthOverflowWarning -f $tableWidthRenderPct) -IsWarning;
                        }
                    }
                    [ref] $null = $tblW.SetAttribute('w', $xmlnsMain, $tableWidthRenderPct * 50);
                }
                $spacing = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 72);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 72);
                #$tblLook = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLook', $xmlnsMain));
                #[ref] $null = $tblLook.SetAttribute('val', $xmlnsMain, '04A0');
                #[ref] $null = $tblLook.SetAttribute('firstRow', $xmlnsMain, 1);
                ## <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
                #$tblStyle = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyle', $xmlnsMain));
                #[ref] $null = $tblStyle.SetAttribute('val', $xmlnsMain, $Table.Style);
                if ($tableStyle.BorderWidth -gt 0) {
                    $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain));
                    foreach ($border in @('top','bottom','start','end','insideH','insideV')) {
                        $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain));
                        [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips $tableStyle.BorderWidth));
                        [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single');
                        [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor));
                    }
                }
                $tblCellMar = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblCellMar', $xmlnsMain));
                $top = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'top', $xmlnsMain));
                [ref] $null = $top.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingTop));
                [ref] $null = $top.SetAttribute('type', $xmlnsMain, 'dxa');
                $left = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'start', $xmlnsMain));
                [ref] $null = $left.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingLeft));
                [ref] $null = $left.SetAttribute('type', $xmlnsMain, 'dxa');
                $bottom = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain));
                [ref] $null = $bottom.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingBottom));
                [ref] $null = $bottom.SetAttribute('type', $xmlnsMain, 'dxa');
                $right = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'end', $xmlnsMain));
                [ref] $null = $right.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingRight));
                [ref] $null = $right.SetAttribute('type', $xmlnsMain, 'dxa');
                $tblGrid = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblGrid', $xmlnsMain));
                $columnCount = $Table.Columns.Count;
                if ($Table.List) {
                    $columnCount = 2;
                }
                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                    [ref] $null = $tblGrid.AppendChild($XmlDocument.CreateElement('w', 'gridCol', $xmlnsMain));
                }
                return $tbl;
            } #end process
        } #end function GetWordTable
        function OutWordTable {
        <#
            .SYNOPSIS
                Output formatted Word table.
            .NOTES
                Specifies that the current row should be repeated at the top each new page on which the table is displayed. E.g, <w:tblHeader />.
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,
                ## Root element to append the table(s) to. List view will create multiple tables
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                [System.Xml.XmlElement] $Element,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tableStyle = $Document.TableStyles[$Table.Style];
                $headerStyle = $Document.Styles[$tableStyle.HeaderStyle];
                if ($Table.List) {
                    for ($r = 0; $r -lt $Table.Rows.Count; $r++) {
                        $row = $Table.Rows[$r];
                        if ($r -gt 0) {
                            ## Add a space between each table as Word renders them together..
                            [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                        }
                        ## Create <tr><tc></tc></tr> for each property
                        $tbl = $Element.AppendChild((GetWordTable -Table $Table -XmlDocument $XmlDocument));
                        $properties = @($row.PSObject.Properties);
                        for ($i = 0; $i -lt $properties.Count; $i++) {
                            $propertyName = $properties[$i].Name;
                            ## Ignore __Style properties
                            if (-not $propertyName.EndsWith('__Style')) {
                                $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain));
                                $tc1 = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                                $tcPr1 = $tc1.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain));
                                if ($null -ne $Table.ColumnWidths) {
                                    ## TODO: Refactor out
                                    [ref] $null = ConvertMmToTwips -Millimeter $Table.ColumnWidths[0];
                                    $tcW1 = $tcPr1.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain));
                                    [ref] $null = $tcW1.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[0] * 50);
                                    [ref] $null = $tcW1.SetAttribute('type', $xmlnsMain, 'pct');
                                }
                                if ($headerStyle.BackgroundColor) {
                                    [ref] $null = $tc1.AppendChild((GetWordTableStyleCellPr -Style $headerStyle -XmlDocument $XmlDocument));
                                }
                                $p1 = $tc1.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                                $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                                $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                                [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, $tableStyle.HeaderStyle);
                                $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($propertyName));
                                $tc2 = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                                $tcPr2 = $tc2.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain));
                                if ($null -ne $Table.ColumnWidths) {
                                    ## TODO: Refactor out
                                    $tcW2 = $tcPr2.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain));
                                    [ref] $null = $tcW2.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[1] * 50);
                                    [ref] $null = $tcW2.SetAttribute('type', $xmlnsMain, 'pct');
                                }
                                $p2 = $tc2.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                                $cellPropertyStyle = '{0}__Style' -f $propertyName;
                                if ($row.PSObject.Properties[$cellPropertyStyle]) {
                                    if (-not (Test-Path -Path Variable:\cellStyle)) {
                                        $cellStyle = $Document.Styles[$row.$cellPropertyStyle];
                                    }
                                    elseif ($cellStyle.Id -ne $row.$cellPropertyStyle) {
                                        ## Retrieve the style if we don't already have it
                                        $cellStyle = $Document.Styles[$row.$cellPropertyStyle];
                                    }
                                    if ($cellStyle.BackgroundColor) {
                                        [ref] $null = $tc2.AppendChild((GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument));
                                    }
                                    if ($row.$cellPropertyStyle) {
                                        $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                                        $pStyle2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                                        [ref] $null = $pStyle2.SetAttribute('val', $xmlnsMain, $row.$cellPropertyStyle);
                                    }
                                }
                                ## Create a separate run for each line/break
                                $lines = $row.($propertyName).ToString() -split [System.Environment]::NewLine;
                                for ($l = 0; $l -lt $lines.Count; $l++) {
                                    $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                    $t2 = $r2.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                    [ref] $null = $t2.AppendChild($XmlDocument.CreateTextNode($lines[$l]));
                                    if ($l -lt ($lines.Count -1)) {
                                        ## Don't add a line break to the last line/break
                                        $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                        $t3 = $r3.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                        [ref] $null = $t3.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                                    }
                                } #end foreach line break
                            }
                        } #end for each property
                     } #end foreach row
                } #end if Table.List
                else {
                    $tbl = $Element.AppendChild((GetWordTable -Table $Table -XmlDocument $XmlDocument));
                    $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain));
                    $trPr = $tr.AppendChild($XmlDocument.CreateElement('w', 'trPr', $xmlnsMain));
                    [ref] $rblHeader = $trPr.AppendChild($XmlDocument.CreateElement('w', 'tblHeader', $xmlnsMain)); ## Flow headers across pages
                    for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
                        $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                        if ($headerStyle.BackgroundColor) {
                            $tcPr = $tc.AppendChild((GetWordTableStyleCellPr -Style $headerStyle -XmlDocument $XmlDocument));
                        }
                        else {
                            $tcPr = $tc.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain));
                        }
                        $tcW = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain));
                        if (($null -ne $Table.ColumnWidths) -and ($null -ne $Table.ColumnWidths[$i])) {
                            [ref] $null = $tcW.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[$i] * 50);
                            [ref] $null = $tcW.SetAttribute('type', $xmlnsMain, 'pct');
                        }
                        else {
                            [ref] $null = $tcW.SetAttribute('w', $xmlnsMain, 0);
                            [ref] $null = $tcW.SetAttribute('type', $xmlnsMain, 'auto');
                        }
                        $p = $tc.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                        $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                        [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $tableStyle.HeaderStyle);
                        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                        $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                        [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($Table.Columns[$i]));
                    } #end for Table.Columns
                    $isAlternatingRow = $false;
                    foreach ($row in $Table.Rows) {
                        $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain));
                        foreach ($propertyName in $Table.Columns) {
                            $cellPropertyStyle = '{0}__Style' -f $propertyName;
                            if ($row.PSObject.Properties[$cellPropertyStyle]) {
                                ## Cell style overrides row/default styles
                                $cellStyleName = $row.$cellPropertyStyle;
                            }
                            elseif (-not [System.String]::IsNullOrEmpty($row.__Style)) {
                                ## Row style overrides default style
                                $cellStyleName = $row.__Style;
                            }
                            else {
                                ## Use the table row/alternating style..
                                $cellStyleName = $tableStyle.RowStyle;
                                if ($isAlternatingRow) {
                                    $cellStyleName = $tableStyle.AlternateRowStyle;
                                }
                            }
                            if (-not (Test-Path -Path Variable:\cellStyle)) {
                                $cellStyle = $Document.Styles[$cellStyleName];
                            }
                            elseif ($cellStyle.Id -ne $cellStyleName) {
                                ## Retrieve the style if we don't already have it
                                $cellStyle = $Document.Styles[$cellStyleName];
                            }
                            $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain));
                            if ($cellStyle.BackgroundColor) {
                                [ref] $null = $tc.AppendChild((GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument));
                            }
                            $p = $tc.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                            $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                            [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $cellStyleName);
                            ## Create a separate run for each line/break
                            $lines = $row.($propertyName).ToString() -split [System.Environment]::NewLine;
                            for ($l = 0; $l -lt $lines.Count; $l++) {
                                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($lines[$l]));
                                if ($l -lt ($lines.Count -1)) {
                                    ## Don't add a line break to the last line/break
                                    $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                    $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                    [ref] $null = $t.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                                }
                            } #end foreach line break
                        } #end foreach property
                        $isAlternatingRow = !$isAlternatingRow;
                    } #end foreach row
                } #end if not Table.List
            } #end process
        } #end function OutWordTable
        function OutWordTOC {
        <#
            .SYNOPSIS
                 Output formatted Word table of contents.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $sdt = $XmlDocument.CreateElement('w', 'sdt', $xmlnsMain);
                $sdtPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'sdtPr', $xmlnsMain));
                $docPartObj = $sdtPr.AppendChild($XmlDocument.CreateElement('w', 'docPartObj', $xmlnsMain));
                $docObjectGallery = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartGallery', $xmlnsMain));
                [ref] $null = $docObjectGallery.SetAttribute('val', $xmlnsMain, 'Table of Contents');
                [ref] $null = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartUnique', $xmlnsMain));
                [ref] $null = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdEndPr', $xmlnsMain));
                $sdtContent = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdContent', $xmlnsMain));
                $p1 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain));
                [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, 'TOC');
                $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($TOC.Name));
                $p2 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                $tabs2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'tabs', $xmlnsMain));
                $tab2 = $tabs2.AppendChild($XmlDocument.CreateElement('w', 'tab', $xmlnsMain));
                [ref] $null = $tab2.SetAttribute('val', $xmlnsMain, 'right');
                [ref] $null = $tab2.SetAttribute('leader', $xmlnsMain, 'dot');
                [ref] $null = $tab2.SetAttribute('pos', $xmlnsMain, '9016'); #10790?!
                $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                ##TODO: Refactor duplicate code
                $fldChar1 = $r2.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain));
                [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlnsMain, 'begin');
                $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $instrText = $r3.AppendChild($XmlDocument.CreateElement('w', 'instrText', $xmlnsMain));
                [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve');
                [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' TOC \o "1-3" \h \z \u '));
                $r4 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                $fldChar2 = $r4.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain));
                [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlnsMain, 'separate');
                $p3 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $r5 = $p3.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                #$rPr3 = $r3.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain));
                $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain));
                [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlnsMain, 'end');
                return $sdt;
            } #end process
        } #end function OutWordTOC
        function OutWordBlankLine {
        <#
            .SYNOPSIS
                Output formatted Word xml blank line (paragraph).
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $BlankLine,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument,
                [Parameter(Mandatory)]
                [System.Xml.XmlElement] $Element
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
                    [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                }
            }
        } #end function OutWordLineBreak
        function GetWordStyle {
        <#
            .SYNOPSIS
                Generates Word Xml style element from a PScribo document style.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Style,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument,
                [Parameter(Mandatory)]
                [ValidateSet('Paragraph','Character')]
                [System.String] $Type
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                if ($Type -eq 'Paragraph') {
                    $styleId = $Style.Id;
                    $styleName = $Style.Name;
                    $linkId = '{0}Char' -f $Style.Id;
                }
                else {
                    $styleId = '{0}Char' -f $Style.Id;
                    $styleName = '{0} Char' -f $Style.Name;
                    $linkId = $Style.Id;
                }
                $documentStyle = $XmlDocument.CreateElement('w', 'style', $xmlnsMain);
                [ref] $null = $documentStyle.SetAttribute('type', $xmlnsMain, $Type.ToLower());
                if ($Style.Id -eq $Document.DefaultStyle) {
                    ## Set as default style
                    [ref] $null = $documentStyle.SetAttribute('default', $xmlnsMain, 1);
                    $uiPriority = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'uiPriority', $xmlnsMain));
                    [ref] $null = $uiPriority.SetAttribute('val', $xmlnsMain, 1);
                }
                elseif ($Style.Hidden -eq $true) {
                    ## Semi hide style (headers and footers etc)
                    [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain));
                }
                elseif (($document.TableStyles.Values | ForEach-Object { $_.HeaderStyle; $_.RowStyle; $_.AlternateRowStyle; }) -contains $Style.Id) {
                    ## Semi hide styles behind table styles (except default style!)
                    [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain));
                }
                [ref] $null = $documentStyle.SetAttribute('styleId', $xmlnsMain, $styleId);
                $documentStyleName = $documentStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlnsMain));
                [ref] $null = $documentStyleName.SetAttribute('val', $xmlnsMain, $styleName);
                $basedOn = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'basedOn', $xmlnsMain));
                [ref] $null = $basedOn.SetAttribute('val', $XmlnsMain, 'Normal');
                $link = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'link', $xmlnsMain));
                [ref] $null = $link.SetAttribute('val', $XmlnsMain, $linkId);
                $next = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'next', $xmlnsMain));
                [ref] $null = $next.SetAttribute('val', $xmlnsMain, 'Normal');
                [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlnsMain));
                $pPr = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain));
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain));
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0);
                ## Set the <w:jc> (justification) element
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain));
                if ($Style.Align.ToLower() -eq 'justify') {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute');
                }
                else {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower());
                }
                if ($Style.BackgroundColor) {
                    $shd = $pPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain));
                    [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear');
                    [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto');
                    [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor));
                }
                [ref] $null = $documentStyle.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument));
                return $documentStyle;
            } #end process
        } #end function GetWordStyle
        function GetWordTableStyle {
        <#
            .SYNOPSIS
                Generates Word Xml table style element from a PScribo document table style.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TableStyle,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $style = $XmlDocument.CreateElement('w', 'style', $xmlnsMain);
                [ref] $null = $style.SetAttribute('type', $xmlnsMain, 'table');
                [ref] $null = $style.SetAttribute('styleId', $xmlnsMain, $TableStyle.Id);
                $name = $style.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain));
                [ref] $null = $name.SetAttribute('val', $xmlnsMain, $TableStyle.Id);
                $tblPr = $style.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain));
                $tblStyleRowBandSize = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyleRowBandSize', $xmlnsMain));
                [ref] $null = $tblStyleRowBandSize.SetAttribute('val', $xmlnsMain, 1);
                if ($tableStyle.BorderWidth -gt 0) {
                    $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain));
                    foreach ($border in @('top','bottom','start','end','insideH','insideV')) {
                        $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain));
                        [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips $tableStyle.BorderWidth));
                        [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single');
                        [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor));
                    }
                }
                [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$TableStyle.HeaderStyle] -Type Header -XmlDocument $XmlDocument));
                [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$TableStyle.RowStyle] -Type Row -XmlDocument $XmlDocument));
                [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$TableStyle.AlternateRowStyle] -Type AlternateRow -XmlDocument $XmlDocument));
                return $style;
            }
        } #end function GetWordTableStyle
        function GetWordStyleParagraphPr {
        <#
            .SYNOPSIS
                Generates Word paragraph (pPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $pPr = $XmlDocument.CreateElement('w', 'pPr', $xmlnsMain);
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0);
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0);
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain));
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain));
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain));
                if ($Style.Align.ToLower() -eq 'justify') { [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute'); }
                else { [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower()); }
                return $pPr;
            } #end process
        } #end function GetWordTableCellPr
        function GetWordStyleRunPrColor {
        <#
            .SYNOPSIS
                Generates Word run (rPr) text colour formatting property only.
            .NOTES
                This is only required to override the text colour in table rows/headers
                as I can't get this (yet) applied via the table style?
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain);
                $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain));
                [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color));
                return $rPr;
            }
        } #end function GetWordStyleRunPrColor
        function GetWordStyleRunPr {
        <#
            .SYNOPSIS
                Generates Word run (rPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain);
                $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain));
                [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Style.Font[0]);
                [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Style.Font[0]);
                if ($Style.Bold) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain));
                }
                if ($Style.Underline) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain));
                }
                if ($Style.Italic) {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain));
                }
                $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain));
                [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color));
                $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain));
                [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Style.Size * 2);
                return $rPr;
            } #end process
        } #end function GetWordStyleRunPr
        function GetWordTableStyleCellPr {
        <#
            .SYNOPSIS
                Generates Word table cell (tcPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tcPr = $XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain);
                if ($Style.BackgroundColor) {
                    $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain));
                    [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear');
                    [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto');
                    [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor));
                }
                return $tcPr;
            } #end process
        } #end function GetWordTableCellPr
        function GetWordTableStylePr {
        <#
            .SYNOPSIS
                Generates Word table style (tblStylePr) formatting properties for specified table style type
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Object] $Style,
                [Parameter(Mandatory)]
                [ValidateSet('Header','Row','AlternateRow')]
                [System.String] $Type,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $tblStylePr = $XmlDocument.CreateElement('w', 'tblStylePr', $xmlnsMain);
                [ref] $null = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain));
                switch ($Type) {
                    'Header' { $tblStylePrType = 'firstRow'; }
                    'Row' { $tblStylePrType = 'band2Horz'; }
                    'AlternateRow' { $tblStylePrType = 'band1Horz'; }
                }
                [ref] $null = $tblStylePr.SetAttribute('type', $xmlnsMain, $tblStylePrType);
                [ref] $null = $tblStylePr.AppendChild((GetWordStyleParagraphPr -Style $Style -XmlDocument $XmlDocument));
                [ref] $null = $tblStylePr.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument));
                [ref] $null = $tblStylePr.AppendChild((GetWordTableStyleCellPr -Style $Style -XmlDocument $XmlDocument));
                return $tblStylePr;
            } #end process
        } #end function GetWordTableStylePr
        function GetWordSectionPr {
        <#
            .SYNOPSIS
                Outputs Office Open XML section element to set page size and margins.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param (
                [Parameter(Mandatory)]
                [System.Single] $PageWidth,
                [Parameter(Mandatory)]
                [System.Single] $PageHeight,
                [Parameter(Mandatory)]
                [System.Single] $PageMarginTop,
                [Parameter(Mandatory)]
                [System.Single] $PageMarginLeft,
                [Parameter(Mandatory)]
                [System.Single] $PageMarginBottom,
                [Parameter(Mandatory)]
                [System.Single] $PageMarginRight,
                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $sectPr = $XmlDocument.CreateElement('w', 'sectPr', $xmlnsMain);
                $pgSz = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgSz', $xmlnsMain));
                [ref] $null = $pgSz.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageWidth));
                [ref] $null = $pgSz.SetAttribute('h', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageHeight));
                [ref] $null = $pgSz.SetAttribute('orient', $xmlnsMain, 'portrait');
                $pgMar = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgMar', $xmlnsMain));
                [ref] $null = $pgMar.SetAttribute('top', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginTop));
                [ref] $null = $pgMar.SetAttribute('bottom', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginBottom));
                [ref] $null = $pgMar.SetAttribute('left', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginLeft));
                [ref] $null = $pgMar.SetAttribute('right', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginRight));
                return $sectPr;
            } #end process
        } #end GetWordSectionPr
        function OutWordStylesDocument {
        <#
            .SYNOPSIS
                Outputs Office Open XML style document
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlDocument])]
            param (
                ## PScribo document styles
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Collections.Hashtable] $Styles,
                ## PScribo document tables styles
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Collections.Hashtable] $TableStyles
            )
            process {
                ## Create the Style.xml document
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument';
                [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
                $documentStyles = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'styles', $xmlnsMain));
                ## Create default style
                $defaultStyle = $documentStyles.AppendChild($xmlDocument.CreateElement('w', 'style', $xmlnsMain));
                [ref] $null = $defaultStyle.SetAttribute('type', $xmlnsMain, 'paragraph');
                [ref] $null = $defaultStyle.SetAttribute('default', $xmlnsMain, '1');
                [ref] $null = $defaultStyle.SetAttribute('styleId', $xmlnsMain, 'Normal');
                $defaultStyleName = $defaultStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlnsMain));
                [ref] $null = $defaultStyleName.SetAttribute('val', $xmlnsMain, 'Normal');
                [ref] $null = $defaultStyle.AppendChild($xmlDocument.CreateElement('w', 'qFormat', $xmlnsMain));
                foreach ($style in $Styles.Values) {
                    $documentParagraphStyle = GetWordStyle -Style $style -XmlDocument $xmlDocument -Type Paragraph;
                    [ref] $null = $documentStyles.AppendChild($documentParagraphStyle);
                    $documentCharacterStyle = GetWordStyle -Style $style -XmlDocument $xmlDocument -Type Character;
                    [ref] $null = $documentStyles.AppendChild($documentCharacterStyle);
                }
                foreach ($tableStyle in $TableStyles.Values) {
                    $documentTableStyle = GetWordTableStyle -TableStyle $tableStyle -XmlDocument $xmlDocument;
                    [ref] $null = $documentStyles.AppendChild($documentTableStyle);
                }
                return $xmlDocument;
            } #end process
        } #end function OutWordStyleDocument
        function OutWordSettingsDocument {
        <#
            .SYNOPSIS
                Outputs Office Open XML settings document
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlDocument])]
            param (
                [Parameter()]
                [System.Management.Automation.SwitchParameter] $UpdateFields
            )
            process {
                ## Create the Style.xml document
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                # <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                # xmlns:o="urn:schemas-microsoft-com:office:office"
                # xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                # xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                # xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"
                # xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                # xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                # xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                # xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
                # mc:Ignorable="w14 w15">
                $settingsDocument = New-Object -TypeName 'System.Xml.XmlDocument';
                [ref] $null = $settingsDocument.AppendChild($settingsDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
                $settings = $settingsDocument.AppendChild($settingsDocument.CreateElement('w', 'settings', $xmlnsMain));
                ## Set compatibility mode to Word 2013
                $compat = $settings.AppendChild($settingsDocument.CreateElement('w', 'compat', $xmlnsMain));
                $compatSetting = $compat.AppendChild($settingsDocument.CreateElement('w', 'compatSetting', $xmlnsMain));
                [ref] $null = $compatSetting.SetAttribute('name', $xmlnsMain, 'compatibilityMode');
                [ref] $null = $compatSetting.SetAttribute('uri', $xmlnsMain, 'http://schemas.microsoft.com/office/word');
                [ref] $null = $compatSetting.SetAttribute('val', $xmlnsMain, 15);
                if ($UpdateFields) {
                    $wupdateFields = $settings.AppendChild($settingsDocument.CreateElement('w', 'updateFields', $xmlnsMain));
                    [ref] $null = $wupdateFields.SetAttribute('val', $xmlnsMain, 'true');
                }
                return $settingsDocument;
            } #end process
        } #end function OutWordSettingsDocument
        #endregion OutWord Private Functions
    }
    process {
        $stopwatch = [Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument';
        [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
        $documentXml = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'document', $xmlnsMain));
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:xml', 'http://www.w3.org/XML/1998/namespace');
        $body = $documentXml.AppendChild($xmlDocument.CreateElement('w', 'body', $xmlnsMain));
        ## Setup the document page size/margins
        $sectionPrParams = @{
            PageHeight = $Document.Options['PageHeight']; PageWidth = $Document.Options['PageWidth'];
            PageMarginTop = $Document.Options['MarginTop']; PageMarginBottom = $Document.Options['MarginBottom'];
            PageMarginLeft = $Document.Options['MarginLeft']; PageMarginRight = $Document.Options['MarginRight'];
        }
        [ref] $null = $body.AppendChild((GetWordSectionPr @sectionPrParams -XmlDocument $xmlDocument));
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { $s | OutWordSection -RootElement $body -XmlDocument $xmlDocument; }
                'PScribo.Paragraph' { [ref] $null = $body.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $xmlDocument)); }
                'PScribo.PageBreak' { [ref] $null = $body.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $xmlDocument)); }
                'PScribo.LineBreak' { [ref] $null = $body.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $xmlDocument)); }
                'PScribo.Table' { OutWordTable -Table $s -XmlDocument $xmlDocument -Element $body; }
                'PScribo.TOC' { [ref] $null = $body.AppendChild((OutWordTOC -TOC $s -XmlDocument $xmlDocument)); }
                'PScribo.BlankLine' { OutWordBlankLine -BlankLine $s -XmlDocument $xmlDocument -Element $body; }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            } #end switch
        } #end foreach
        ## Generate the Word 'styles.xml' document part
        $stylesXml = OutWordStylesDocument -Styles $Document.Styles -TableStyles $Document.TableStyles;
        ## Generate the Word 'settings.xml' document part
        if (($Document.Properties['TOCs']) -and ($Document.Properties['TOCs'] -gt 0)) {
            ## We have a TOC so flag to update the document when opened
            $settingsXml = OutWordSettingsDocument -UpdateFields;
        }
        else {
            $settingsXml = OutWordSettingsDocument;
        }
        #Convert relative or PSDrive based path to the absolute filesystem path
        $AbsolutePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $destinationPath = Join-Path -Path $AbsolutePath ('{0}.docx' -f $Document.Name);
        if ((-not $PSVersionTable.ContainsKey('PSEdition')) -or ($PSVersionTable.PSEdition -ne 'Core')) {
            ## WindowsBase.dll is not included in Core PowerShell
            Add-Type -AssemblyName WindowsBase;
        }
        try {
            $package = [System.IO.Packaging.Package]::Open($destinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite);
        }
        catch {
            WriteLog -Message ($localized.OpenPackageError -f $destinationPath) -IsWarning;
            throw $_;
        }
        ## Create document.xml part
        $documentUri = New-Object System.Uri('/word/document.xml', [System.UriKind]::Relative);
        WriteLog -Message ($localized.ProcessingDocumentPart -f $documentUri);
        $documentPart = $package.CreatePart($documentUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
        $streamWriter = New-Object System.IO.StreamWriter($documentPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite));
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter);
        WriteLog -Message ($localized.WritingDocumentPart -f $documentUri);
        $xmlDocument.Save($xmlWriter);
        $xmlWriter.Dispose();
        $streamWriter.Close();
        ## Create styles.xml part
        $stylesUri = New-Object System.Uri('/word/styles.xml', [System.UriKind]::Relative);
        WriteLog -Message ($localized.ProcessingDocumentPart -f $stylesUri);
        $stylesPart = $package.CreatePart($stylesUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml');
        $streamWriter = New-Object System.IO.StreamWriter($stylesPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite));
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter);
        WriteLog -Message ($localized.WritingDocumentPart -f $stylesUri);
        $stylesXml.Save($xmlWriter);
        $xmlWriter.Dispose();
        $streamWriter.Close();
        ## Create settings.xml part
        $settingsUri = New-Object System.Uri('/word/settings.xml', [System.UriKind]::Relative);
        WriteLog -Message ($localized.ProcessingDocumentPart -f $settingsUri);
        $settingsPart = $package.CreatePart($settingsUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
        $streamWriter = New-Object System.IO.StreamWriter($settingsPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite));
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter);
        WriteLog -Message ($localized.WritingDocumentPart -f $settingsUri);
        $settingsXml.Save($xmlWriter);
        $xmlWriter.Dispose();
        $streamWriter.Close();
        ## Create the Package relationships
        WriteLog -Message $localized.GeneratingPackageRelationships;
        [ref] $null = $package.CreateRelationship($documentUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'rId1');
        [ref] $null = $documentPart.CreateRelationship($stylesUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'rId1');
        [ref] $null = $documentPart.CreateRelationship($settingsUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'rId2');
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        $package.Flush();
        $package.Close();
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath);
    } #end process
} #end function OutWord

function OutXml {
<#
    .SYNOPSIS
        Xml output plugin for PScribo.
    .DESCRIPTION
        Outputs a xml representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    param (
        ## ThePScribo document object to convert to a xml document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,
        ## Output directory path for the .xml file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,
        ### Hashtable of all plugin supported options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {
        #region OutXml Private Functions
        function OutXmlSection {
        <#
            .SYNOPSIS
                Output formatted Xml section.
        #>
            [CmdletBinding()]
            param (
                ## PScribo document section
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            process {
                $sectionId = ($Section.Id -replace '[^a-z0-9-_\.]','').ToLower();
                $element = $xmlDocument.CreateElement($sectionId);
                [ref] $null = $element.SetAttribute("name", $Section.Name);
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $element.AppendChild((OutXmlSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $element.AppendChild((OutXmlParagraph -Paragraph $s)); }
                        'PScribo.Table' { [ref] $null = $element.AppendChild((OutXmlTable -Table $s)); }
                        'PScribo.PageBreak' { } ## Page breaks are not implemented for Xml output
                        'PScribo.LineBreak' { } ## Line breaks are not implemented for Xml output
                        'PScribo.BlankLine' { } ## Blank lines are not implemented for Xml output
                        'PScribo.TOC' { } ## TOC is not implemented for Xml output
                        Default {
                            WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning;
                        }
                    } #end switch
                } #end foreach
                return $element;
            } #end process
        } #end function outxmlsection
        function OutXmlParagraph {
        <#
            .SYNOPSIS
                Output formatted Xml paragraph.
        #>
            [CmdletBinding()]
            param (
                ## PScribo paragraph object
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            process {
                if (-not ([string]::IsNullOrEmpty($Paragraph.Value))) {
                    ## Value override specified
                    $paragraphId = ($Paragraph.Id -replace '[^a-z0-9-_\.]','').ToLower();
                    $paragraphElement = $xmlDocument.CreateElement($paragraphId);
                    [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($Paragraph.Value));
                } #end if
                elseif ([string]::IsNullOrEmpty($Paragraph.Text)) {
                    ## No Id/Name specified, therefore insert as a comment
                    $paragraphElement = $xmlDocument.CreateComment((' {0} ' -f $Paragraph.Id));
                } #end elseif
                else {
                    ## Create an element with the Id/Name
                    $paragraphId = ($Paragraph.Id -replace '[^a-z0-9-_\.]','').ToLower();
                    $paragraphElement = $xmlDocument.CreateElement($paragraphId);
                    [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($Paragraph.Text));
                } #end else
                return $paragraphElement;
            } #end process
        } #end function outxmlparagraph
        function OutXmlTable {
        <#
            .SYNOPSIS
                Output formatted Xml table.
        #>
            [CmdletBinding()]
            param (
                ## PScribo table object
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {
                $tableId = ($Table.Id -replace '[^a-z0-9-_\.]','').ToLower();
                $tableElement = $element.AppendChild($xmlDocument.CreateElement($tableId));
                [ref] $null = $tableElement.SetAttribute('name', $Table.Name);
                foreach ($row in $Table.Rows) {
                    $groupElement = $tableElement.AppendChild($xmlDocument.CreateElement('group'));
                    foreach ($property in $row.PSObject.Properties) {
                        if (-not ($property.Name).EndsWith('__Style')) {
                            $propertyId = ($property.Name -replace '[^a-z0-9-_\.]','').ToLower();
                            $rowElement = $groupElement.AppendChild($xmlDocument.CreateElement($propertyId));
                            ## Only add the Name attribute if there's a difference
                            if ($property.Name -ne $propertyId) {
                                [ref] $null = $rowElement.SetAttribute('name', $property.Name);
                            }
                            [ref] $null = $rowElement.AppendChild($xmlDocument.CreateTextNode($row.($property.Name)));
                        } #end if
                    } #end foreach property
                } #end foreach row
                return $tableElement;
            } #end process
        } #end outxmltable
        #endregion OutXml Private Functions
    }
    process {
        $pluginName = 'Xml';
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);
        $documentName = $Document.Name;
        $xmlDocument = New-Object -TypeName System.Xml.XmlDocument;
        [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
        $documentId = ($Document.Id -replace '[^a-z0-9-_\.]','').ToLower();
        $element = $xmlDocument.AppendChild($xmlDocument.CreateElement($documentId));
        [ref] $null = $element.SetAttribute("name", $documentName);
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { [ref] $null = $element.AppendChild((OutXmlSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $element.AppendChild((OutXmlParagraph -Paragraph $s)); }
                'PScribo.Table' { [ref] $null = $element.AppendChild((OutXmlTable -Table $s)); }
                'PScribo.PageBreak'{ } ## Page breaks are not implemented for Xml output
                'PScribo.LineBreak' { } ## Line breaks are not implemented for Xml output
                'PScribo.BlankLine' { } ## Blank lines are not implemented for Xml output
                'PScribo.TOC' { } ## TOC is not implemented for Xml output
                Default {
                    WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning;
                }
            } #end switch
        } #end foreach
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        $destinationPath = Join-Path $Path ('{0}.xml' -f $Document.Name);
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        ## Core PowerShell XmlDocument requires a stream
        $streamWriter = New-Object System.IO.StreamWriter($destinationPath, $false);
        $xmlDocument.Save($streamWriter);
        $streamWriter.Close();
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath);
    } #end process
} #end function outxml

#endregion PScribo Bundle v0.7.21.110

# ~~~~~~~~~   end of PScribo bundle   ~~~~~~~~~~

#region Hardware and Software routines

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~< Hardware functions >~~~~~~~~~~~~~~~~~~~~~~~
Function GetComputerWMIInfo ($computername)
{
	# original work by Kees Baggerman, 
	# k.baggerman@myvirtualvision.com
	# 
    # modified 2018-05-25 Sam Jacobs to use PScribo

	#Get Computer info
    if ($computername -eq $Null) { $computername = $env:COMPUTERNAME }

	Write-Verbose "$(Get-Date): Processing WMI information for $($computername)"
	Write-Verbose "$(Get-Date): `tHardware information"
	
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $computername win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($computername)"
		Write-Warning "$(Get-Date): Get-WmiObject win32_computersystem failed for $($computername)"

			Output-Line 0 "Get-WmiObject win32_computersystem failed for $($computername)" "" $Null 0 $False $True
			Output-Line 0 "On $($computername) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			Output-Line 0 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			Output-Line 0 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
			Output-Line 0 "No results Returned for Computer information" "" $Null 0 $False $True
	}

	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

    Section -Style Heading2 "Drive Information" {

	    [bool]$GotDrives = $True
	
	    Try
	    {
		    $Results = Get-WmiObject -computername $computername Win32_LogicalDisk
	    }
	
	    Catch
	    {
		    $Results = $Null
	    }

	    If($? -and $Null -ne $Results)
	    {
		    $drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
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
		    Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($computername)"
		    Write-Warning "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($computername)"
			    Output-Line 0  "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($computername)" "" $Null 0 $False $True
			    Output-Line 0  "$(Get-Date): On $($computername) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			    Output-Line 0  "$(Get-Date): and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			    Output-Line 0  "$(Get-Date): need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	    }
	    Else
	    {
		    Write-Verbose "$(Get-Date): No results Returned for Drive information"

			    Output-Line 0  "$(Get-Date): No results Returned for Drive information" "" $Null 0 $False $True
	    }

    }

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

    Section -Style Heading2 "Processor Information" {

	    [bool]$GotProcessors = $True
	
	    Try
	    {
		    $Results = Get-WmiObject -computername $computername win32_Processor
	    }
	
	    Catch
	    {
		    $Results = $Null
	    }

	    If($? -and $Null -ne $Results)
	    {
		    $Processors = $Results | Select availability, name, description, maxclockspeed, 
		    l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		    $Results = $Null
		    ForEach($processor in $processors)
		    {
			    OutputProcessorItem $processor
		    }
	    }
	    ElseIf(!$?)
	    {
		    Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($computername)"
		    Write-Warning "$(Get-Date): Get-WmiObject win32_Processor failed for $($computername)"

			    Output-Line 0 "$(Get-Date): Get-WmiObject win32_Processor failed for $($computername)" "" $Null 0 $False $True
			    Output-Line 0 "$(Get-Date): On $($computername) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			    Output-Line 0 "$(Get-Date): and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			    Output-Line 0 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	    }
	    Else
	    {
		    Write-Verbose "$(Get-Date): No results Returned for Processor information"
			    Output-Line 0 "$(Get-Date): No results Returned for Processor information" "" $Null 0 $False $True
	    }

    }

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	Section -Style Heading2 "Network Interface(s)" {

	    [bool]$GotNics = $True
	
	    Try
	    {
		    $Results = Get-WmiObject -computername $computername win32_networkadapterconfiguration
	    }
	
	    Catch
	    {
		    $Results = $Null
	    }

	    If($? -and $Null -ne $Results)
	    {
		    $Nics = $Results | Where {$Null -ne $_.ipaddress}
		    $Results = $Null

		    If($Nics -eq $Null ) 
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
					    $ThisNic = Get-WmiObject -computername $computername win32_networkadapter | Where {$_.index -eq $nic.index}
				    }
				
				    Catch 
				    {
					    $ThisNic = $Null
				    }
				
				    If($? -and $Null -ne $ThisNic)
				    {
					    OutputNicItem $Nic $ThisNic
				    }
				    ElseIf(!$?)
				    {
					    Write-Warning "$(Get-Date): Error retrieving NIC information"
					    Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($computername)"
					    Write-Warning "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($computername)"
						    Output-Line 0 "$(Get-Date): Error retrieving NIC information" "" $Null 0 $False $True
						    Output-Line 0 "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($computername)" "" $Null 0 $False $True
						    Output-Line 0 "$(Get-Date): On $($computername) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						    Output-Line 0 "$(Get-Date): and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						    Output-Line 0 "$(Get-Date): need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
				    }
				    Else
				    {
					    Write-Verbose "$(Get-Date): No results Returned for NIC information"
						    Output-Line 0 "$(Get-Date): No results Returned for NIC information" "" $Null 0 $False $True
				    }
			    }
		    }	
	    }
	    ElseIf(!$?)
	    {
		    Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		    Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($computername)"
		    Write-Warning "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($computername)"
			    Output-Line 0 "$(Get-Date): Error retrieving NIC configuration information" "" $Null 0 $False $True
			    Output-Line 0 "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($computername)" "" $Null 0 $False $True
			    Output-Line 0 "$(Get-Date): On $($computername) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			    Output-Line 0 "$(Get-Date): and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			    Output-Line 0 "$(Get-Date): need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	    }
	    Else
	    {
		    Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
			    Output-Line 0 "$(Get-Date): No results Returned for NIC configuration information" "" $Null 0 $False $True
	    }
	}

	Output-Line 0 ""

}

Function OutputComputerItem
{
	Param([object]$Item)

		$rowdata = @()
        $rowdata += @(,("Manufacturer",$Item.Manufacturer))
		$rowdata += @(,('Model',$Item.model))
		$rowdata += @(,('Domain',$Item.domain))
        $physRAM = $Item.totalphysicalram
		$rowdata += @(,('Total Ram',"$($physRAM) GB"))
		$rowdata += @(,('Physical Processors (sockets)',$Item.NumberOfProcessors))
		$rowdata += @(,('Logical Processors (cores w/HT)',$Item.NumberOfLogicalProcessors))

		Output-Table "Computer Information" $rowdata "list" "Heading2"
		Output-Line 0 ""
	
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


		$rowdata = @()
		$rowdata += @(,('Caption',$Drive.caption))
        $driveSize = "$($drive.drivesize) GB"
		$rowdata += @(,('Size',$driveSize))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',$Drive.filesystem))
		}
        $freeSpace = "$($drive.drivefreespace) GB"
		$rowdata += @(,('Free Space',$freeSpace))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',$Drive.volumename))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',$xVolumeDirty))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',$Drive.volumeserialnumber))
		}
		$rowdata += @(,('Drive Type',$xDriveType))

		Output-Table "Drive: $($Drive.caption)" $rowdata "list" "Heading3"
		Output-Line 0 ""	
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


		$rowdata = @()
		$rowdata += @(,('Name',$Processor.name))
		$rowdata += @(,('Description',$Processor.description))
        $maxSpeed = "$($processor.maxclockspeed) MHz"
		$rowdata += @(,('Max Clock Speed',$maxSpeed))
		If($processor.l2cachesize -gt 0)
		{
            $l2cacheSize = "$($processor.l2cachesize) KB"
			$rowdata += @(,('L2 Cache Size',$l2cacheSize))
		}
		If($processor.l3cachesize -gt 0)
		{
            $l3cacheSize = "$($processor.l3cachesize) KB"
			$rowdata += @(,('L3 Cache Size',$l3cacheSize))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',$Processor.numberofcores))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',$Processor.numberoflogicalprocessors))
		}
		$rowdata += @(,('Availability',$xAvailability))

		Output-Table "Processor: $($Processor.name)" $rowdata "list" "Heading3"
		Output-Line 0 ""
	
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	
	$powerMgmt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi | where {$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

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

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
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


		$rowdata = @()
		$rowdata += @(,('Name',$ThisNic.Name))
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',$Nic.description))
		}
		$rowdata += @(,('Connection ID',$ThisNic.NetConnectionID))
        if ($Nic.manufacturer -ne $Null) {
		Write-Verbose "$(Get-Date): Manufacturer: $($Nic.manufacturer)"
		    $rowdata += @(,('Manufacturer',$Nic.manufacturer))
        }
		$rowdata += @(,('Availability',$xAvailability))
		$rowdata += @(,('Allow turn off to save power',$PowerSaving))
		$rowdata += @(,('Physical Address',$Nic.macaddress))
		$rowdata += @(,('IP Address',$xIPAddress[0]))
		Write-Verbose "$(Get-Date): NIC: $($ThisNic.Name)"
		Write-Verbose "$(Get-Date): Description: $($Nic.description)"
		Write-Verbose "$(Get-Date): ConnectionID: $($ThisNic.NetConnectionID)"
		Write-Verbose "$(Get-Date): Availability: $($xAvailability)"
		Write-Verbose "$(Get-Date): Powersaving: $($PowerSaving)"
		Write-Verbose "$(Get-Date): Mac address: $($Nic.macaddress)"
		Write-Verbose "$(Get-Date): IP Address: $($xIPAddress[0])"
		
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
                $strIP = @{$true = "IP Address v6"; $false = "IP Address $($cnt+1)"}[$cnt -eq 1] 
				Write-Verbose "$(Get-Date): $($strIP): $($tmp)"              
				$rowdata += @(,($strIP,$tmp))
			}
		}
        
        if ($Nic.Defaultipgateway -eq $Nul) { $defGW = " " } else { $defGW = $Nic.Defaultipgateway[0] }
		Write-Verbose "$(Get-Date): Def Gateway: $($defGW)"
		$rowdata += @(,('Default Gateway',$defGW))
        if ($xIPSubnet.Count -eq 2) {
            $subNet = "$($xIPSubnet[0]) / $($xIPSubnet[1])"
	    Write-Verbose "$(Get-Date): Subnet Mask: $($subNet)"
            $rowdata += @(,('Subnet Mask',$subNet))
        } else {
		    Write-Verbose "$(Get-Date): Subnet Mask: $($xIPSubnet[0])"
		    $rowdata += @(,('Subnet Mask',$xIPSubnet[0]))
		    $cnt = -1
		    ForEach($tmp in $xIPSubnet)
		    {
			    $cnt++
			    If($cnt -gt 0)
			    {
                    $strSubnet = "Subnet Mask $($cnt+1)"
				    Write-Verbose "$(Get-Date): $($strSubnet): $($tmp)"
				    $rowdata += @(,($strSubnet,$tmp))
			    }
		    }
        }
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',$Nic.dhcpenabled))
			$rowdata += @(,('DHCP Lease Obtained',$dhcpleaseobtaineddate))
			$rowdata += @(,('DHCP Lease Expires',$dhcpleaseexpiresdate))
			$rowdata += @(,('DHCP Server',$Nic.dhcpserver))

			Write-Verbose "$(Get-Date): DHCP Enabled: $($Nic.dhcpenabled))"
			Write-Verbose "$(Get-Date): DHCP Lease Obtained: $($dhcpleaseobtaineddate))"
			Write-Verbose "$(Get-Date): DHCP Lease Expires: $($dhcpleaseexpiresdate))"
			Write-Verbose "$(Get-Date): DHCP Server: $($Nic.dhcpserver))"
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Write-Verbose "$(Get-Date): DNS Domain: $($Nic.dnsdomain)"
			$rowdata += @(,('DNS Domain',$Nic.dnsdomain))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			Write-Verbose "$(Get-Date): DNS Search Suffixes: $($xnicdnsdomainsuffixsearchorder[0])"
			$rowdata += @(,('DNS Search Suffixes',$xnicdnsdomainsuffixsearchorder[0]))
			$cnt = -1
			$tmpVal = ""
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				If($cnt -gt 0)
				{
					$tmpVal += "`r`n"
				}
				$cnt++
				If($cnt -gt 0)
				{
					$tmpVal += $tmp
				}
			}
			If($cnt -gt 0)
			{
				$rowdata += @(,('Domain Suffix Search Order',$tmp))
			}
		}
		$rowdata += @(,('DNS WINS Enabled',$xdnsenabledforwinsresolution))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',$xnicdnsserversearchorder[0]))
			$cnt = -1
			$tmpVal = ""
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				If($cnt -gt 0)
				{
					$tmpVal += "`r`n"
				}
				$cnt++
				If($cnt -gt 0)
				{
					$tmpVal += $tmp
				}
			}
			If($cnt -gt 0)
			{
				$rowdata += @(,('DNS Server Search Order',$tmp))
			}
		}
		$rowdata += @(,('NetBIOS Setting',$xTcpipNetbiosOptions))
		$rowdata += @(,('WINS: Enabled LMHosts',$xwinsenablelmhostslookup))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',$Nic.winshostlookupfile))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',$Nic.winsprimaryserver))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',$Nic.winssecondaryserver))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',$Nic.winsscopeid))
		}

		Output-Table "NIC: $($ThisNic.Name)" $rowdata "list" "Heading3"
		Output-Line 0 ""
	
}

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~< Software functions >~~~~~~~~~~~~~~~~~~~~~~~

Function ConvertTo-ScriptBlock 
{
	#by Jeff Wouters, PowerShell MVP
	Param([string]$string)
	$ScriptBlock = $executioncontext.invokecommand.NewScriptBlock($string)
	Return $ScriptBlock
} 

Function SWExclusions 
{
        # original work by Shaun Ritchie, Jeff Wouters, Webster
        # modified 2018-05-27 Sam Jacobs

        $var = ""
        $Tmp = '$InstalledApps | Where {'
        $Exclusions = @(Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0)
        If($? -and $Exclusions.Count -gt 0)
        {
            ForEach($Exclusion in $Exclusions) 
            {
                $Tmp += "(`$`_.DisplayName -notlike ""$($Exclusion)"") -and "
            }
            $var += $Tmp.Substring(0,($Tmp.Length - 6))
 
            $var += "} | Select-Object DisplayName, DisplayVersion | Sort DisplayName -unique"
        }
        return $var
}

Function InstalledSoftware ($computername)
{
    if ($computername -eq $Null) { $computername = $env:ComputerName }

	#get list of applications installed on server
	# original code by Shaun Ritchie, Jeff Wouters, Webster, Michael B. Smith
	$InstalledApps = @()
	$JustApps = @()

	#Define the variable to hold the location of Currently Installed Programs
	$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

	#Create an instance of the Registry Object and open the HKLM base key
	$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $computername) 

	#Drill down into the Uninstall key using the OpenSubKey Method
	$regkey1=$reg.OpenSubKey($UninstallKey1) 

	#Retrieve an array of string that contain all the subkey names
	If($regkey1 -ne $Null)
	{
		$subkeys1=$regkey1.GetSubKeyNames() 

		#Open each Subkey and use GetValue Method to return the required values for each
		ForEach($key in $subkeys1) 
		{
			$thisKey=$UninstallKey1+"\\"+$key 
			$thisSubKey=$reg.OpenSubKey($thisKey) 
			If(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
			{
				$obj = New-Object PSObject
				$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
				$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
				$InstalledApps += $obj
			}
		}
	}			

	$UninstallKey2="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
	$regkey2=$reg.OpenSubKey($UninstallKey2)
	If($regkey2 -ne $Null)
	{
		$subkeys2=$regkey2.GetSubKeyNames()

		ForEach($key in $subkeys2) 
		{
			$thisKey=$UninstallKey2+"\\"+$key 
			$thisSubKey=$reg.OpenSubKey($thisKey) 
			if(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
			{
				$obj = New-Object PSObject
				$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
				$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
				$InstalledApps += $obj
			}
		}
	}

	$InstalledApps = $InstalledApps | Sort DisplayName

	$tmp1 = SWExclusions
	If($Tmp1 -ne "")
	{
		$Func = ConvertTo-ScriptBlock $tmp1
		$tempapps = Invoke-Command {& $Func}
	}
	Else
	{
		$tempapps = $InstalledApps
	}
	
	$JustApps = $TempApps | Select DisplayName, DisplayVersion | Sort DisplayName -unique

	Write-Verbose "$(Get-Date): `t`tProcessing installed applications for $($computername)"
	
    $apps = @()
    $apps += @(,('Name', 'Version'))
	ForEach($app in $JustApps)
	{
		Write-Verbose "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)"
        if ($app.DisplayVersion -eq $Null) { $app.DisplayVersion = ' ' }
        $apps += @(,($app.DisplayName, $app.DisplayVersion))
	}

	Output-Table "Applications" $apps "table" "Heading2"
	Output-Line 0 ""

}

Function CitrixServices ($computername)
{				
    if ($computername -eq $Null) { $computername = $env:ComputerName }

		Write-Verbose "$(Get-Date): `t`tProcessing Citrix services for server $($computername) by calling Get-Service"

		Try
		{
			$Services = @(Get-WMIObject Win32_Service -ComputerName $computername -EA 0 | Where {$_.DisplayName -like "*Citrix*"} | Sort DisplayName)
		}
		
		Catch
		{
			$Services = $Null
		}

		If($? -and $Services.Count -gt 0)
		{
			[int]$NumServices = $Services.count

			Write-Verbose "$(Get-Date): `t`t $NumServices Services found"
            $svcdata = @()
			$svcdata += @(,('Name', 'State', 'StartMode'))
			ForEach($Service in $Services) 
			{
				Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";
                $svcdata += @(,($Service.DisplayName, $Service.State, $Service.StartMode))
			}
		    Output-Table "Citrix Services" $svcdata "table" "Heading2"
		    Output-Line 0 ""

		}
		ElseIf(!$?)
		{
			Write-Warning "$(Get-Date): No services were retrieved."
		}
		Else
		{
			Write-Warning "$(Get-Date): Services retrieval was successful but no services were returned."
		}

}
#~~~~~~~~~~~~~~~~< End of hardware and software functions >~~~~~~~~~~~~~~~~

#endregion Hardware and Software routines

#region SFDoc 
Function Get-ReplicationObject ($mem)
{
	if ($mem -eq $env:computername)
		{ return Get-ItemProperty -Path HKLM:\Software\Citrix\DeliveryServices\ConfigurationReplication } 
	else
		{ return Invoke-Command -ComputerName $mem -Command {Get-ItemProperty -Path HKLM:\Software\Citrix\DeliveryServices\ConfigurationReplication}  }
}

Function Get-ReplicationStatus ($mem)
{
    # get local or remote replication information
    $repObject = Get-ReplicationObject $mem
    $today = Get-Date
    $repTime = Get-Date $repObject.LastEndTime
    $strTime = Get-Date $repObject.LastEndTime -f t

    if ($reptime.Date -eq $today.Date)
        { $strDate = "today" }
    else
        { $strDate = "on $(Get-Date $repObject.LastEndTime -f d)" }
    $sourceServer = $repObject.LastSourceServer
    if ($mem -eq $sourceServer)
    {
        $repStatus = "Propagated changes $($strDate) at $($strTime). "
        if ($repObject.LastUpdateStatus -eq "Complete") 
        { $repStatus += "All servers are in sync." }
        else
        { $repStatus += $repObject.LastErrorMessage }
    }
    else
    {
        $repStatus = "Synchronized settings with $($sourceServer) $($strDate) at $($strTime). "
        if ($repObject.LastUpdateStatus -ne "Complete") 
        { $repStatus += $repObject.LastErrorMessage }
    }
    return $repStatus
}

Function Import-STFModules()
{
    Write-Host "$(Get-Date): Loading StoreFront modules ..."
    Add-PSSnapin Citrix.DeliveryServices.*
}

# Redirect output to Pscribo
# Useful if we need to switch back to Write-Host for debugging

Function Output-Line ($line, $tabs)
{
	PScribo-Line $line $tabs
}

Function Output-Table ($header,$rows,$listOrTable, $heading)
{
	Pscribo-Table $header $rows $listOrTable $heading
}

Function PScribo-Table ($title,$rows,$listOrTable, $heading)
{
    switch ($heading) {
        "Heading1" {$tabs = 0}
        "Heading2" {$tabs = 1}
        "Heading3" {$tabs = 2}
    }

	Section -Style $heading $title  {
		if ($listOrTable -eq "table")
		{
			# headings are in the first row
			$headers = $rows[0]
			# create an array of custom objects from the remaining rows
			$objs = @()
			for($idx=1;$idx -lt $rows.Count; $idx++) {
				$row = $rows[$idx]
				$obj = New-Object System.Object
				for($idx2=0;$idx2 -lt $row.Count; $idx2++) {
					$obj | Add-Member -type NoteProperty -name $headers[$idx2] -Value $row[$idx2]
				}
				$objs += $obj
			}
            # would like to be able to set tabs for the table (and its header!) at some point
            # $objs | Table -Columns $headers -Width 0 -Tabs $tabs
			$objs | Table -Columns $headers -Width 0
		}
		else
		{
			# list tables are 2 columns - convert to custom object as above
            $headers = @()
			$objs = @()
            $obj = New-Object System.Object
			for($idx=0;$idx -lt $rows.Count; $idx++) {
				$row = $rows[$idx]
                $headers += $row[0]
				$obj | Add-Member -type NoteProperty -name $row[0] -Value $row[1]
			}
		    $objs += $obj
            # would like to be able to set tabs for the table (and its header!) at some point
			# $objs | Table -Columns $headers -List -Width 0 -Tabs $tabs
            $objs | Table -Columns $headers -List -Width 0 
		}
	}

}

Function PScribo-Line ($line, $tabs)
{
   if ($tabs -eq $null) { $tabs = 0 }

   if ($line -eq "")
   { BlankLine; }
   else
   { Paragraph $line -Tabs $tabs}
}

Function Translate-Enabled ([boolean] $condition) {
    if ($condition -eq $True) {
	return "Enabled"
    } Else {
	return "Disabled"
    }
}

Function Translate-Method($methodKey) {
	Switch($methodKey) {
		"ExplicitForms" 	{$methodValue = "User name and password"}
		"Forms-Saml" 	{$methodValue = "SAML Authentication"}
		"IntegratedWindows"	{$methodValue = "Domain pass-through"}
		"CitrixAGBasic"		{$methodValue = "Pass-through from NetScaler"}
		"HttpBasic"		{$methodValue = "HTTP Basic"}
		"Certificate"		{$methodValue = "Smart card"}
		Default			{$methodValue = ""}
	}
	return $methodValue
}

Function Translate-NSVersion($versionKey) {
	Switch($versionKey) {
		"Version10_0_69_4" 	{$versionValue = "10.0 (Build 69.4) or later"}
		"Version9x"		{$versionValue = "9.x"}
		"Version5x"		{$versionValue = "5.x"}
		Default			{$versionValue = $versionKey}
	}
	return $versionValue
}

Function Translate-HTML5Deployment($HTML5Key) {
	Switch($HTML5Key) {
		"Fallback" 	{$HTML5Value = "Use Receiver for HTML5 if local install fails"}
		"Always"	{$HTML5Value = "Always use Receiver for HTML5"}
		"Off"		{$HTML5Value = "Citrix Receiver installed locally"}
		Default		{$HTML5Value = $HTML5Key}
	}
	return $HTML5Value
}

Function Translate-LogonType($logonKey) {
	Switch($logonKey) {
		"DomainAndRSA" 	{$logonValue = "Domain and security token"}
		"Domain"	{$logonValue = "Domain"}
		"RSA"		{$logonValue = "Security token"}
		"SMS"		{$logonValue = "SMS authentication"}
		"SmartCard"	{$logonValue = "Smart card"}
		"None"		{$logonValue = "None"}
		Default		{$logonValue = $logonKey}
	}
	return $logonValue
}

Function Translate-PasswordOptions($pwKey) {
	Switch($pwKey) {
		"Always" 	{$pwValue = "At any time"}
		"ExpiredOnly" 	{$pwValue = "When expired"}
		"Never" 	{$pwValue = "Never"}
		Default		{$pwValue = $pwKey}
	}
	return $pwValue
}

Function Translate-PWReminder($remindKey)
{
	Switch($remindKey) {
		"Windows" 	{$remindValue = "Use settings in AD group policy"}
		"Custom" 	{$remindValue = "$($explicitOptions.PasswordExpiryWarningPeriod) days before expiration"}
		"Never" 	{$remindValue = "Do not remind"}
		Default		{$remindValue = $remindKey}
	}
	return $remindValue
}

Function Translate-RemoteAccess($remoteKey) {
	Switch($remoteKey) {
		"StoresOnly" 	{$remoteValue = "Enabled (No VPN Tunnel)"}
		"FullVPN"   {$remoteValue = "Enabled (Full VPN)"}
		Default		{$remoteValue = "Disabled"}
	}
	return $remoteValue
}

Function Translate-YesNo ([boolean] $condition) {
    if ($condition -eq $True) {
	return "Yes"
    } Else {
	return "No"
    }
}

# ~~~~~~ mainline of script begins here! ~~~~~~~~

Try {
  Write-Verbose "$(Get-Date): Checking for valid StoreFront installation"
  $dsInstallProp = Get-ItemProperty -Path HKLM:\SOFTWARE\Citrix\DeliveryServicesManagement -Name InstallDir -EA 0
  if ($dsInstallProp -ne $Null) {
	Import-STFModules
  }
} Catch {
  $dsInstallProp = $Null
}

if ($dsInstallProp -eq $Null) {
  Write-Verbose "$(Get-Date): Server does not have StoreFront installed ..."
  Write-Host "$(Get-Date): Server does not have StoreFront installed."  -ForegroundColor Red -BackgroundColor Black
  Write-Host "$(Get-Date): Aborting script."
  Return
}

$stf = Get-STFDeployment
if (-not $stf) {
   # SF not installed ... message already displayed, so just exit
   exit
}

Write-Host "$(Get-Date): Processing ..."

$document = Document $Script:OutputFile {
    <#  Set the page size to Letter with 0.5inch margins #>
    DocumentOption -PageSize Letter -Margin 36;
    BlankLine -Count 20;
    Paragraph $Script:Title -Style Title;
    BlankLine -Count 15;
    Paragraph "Created by: $($Script:Author)"
    $today = date -Format 'MMMM dd, yyyy'
    Paragraph "$($today)"
    PageBreak;
    TOC -Name 'Table of Contents';
    PageBreak;

    # used in multiple sections
    $serverGroup = Get-STFServerGroup
    $baseURL = $servergroup.HostBaseUrl.AbsoluteUri
    $stores = @(Get-STFStoreService)

    Section -Style Heading1 'Server Group' {
	
		# server group info
		##$serverGroup = Get-STFServerGroup
		##$baseURL = $servergroup.HostBaseUrl.AbsoluteUri
		$memberCount = ($serverGroup.ClusterMembers).Count
		$members = @($serverGroup.ClusterMembers).Server
		$protocol = $servergroup.HostBaseUrl.Scheme
		$group = @()
		$group += @(,("Base URL:", "$($baseUrl)"))
		$group += @(,("Number of Servers:", $memberCount))
		if ($memberCount -gt 1)
		{
			# last source server
			$lastPropSource = (Get-ItemProperty -Path HKLM:\Software\Citrix\DeliveryServices\ConfigurationReplication).LastSourceServer
			$group += @(,("Configuration: ", "Last propagated from $($lastPropSource)"))
		}
		$group += @(,("StoreFront using: ", $protocol))

		# StoreFront version info
		$stfVersion = (Get-STFVersion) -join "."
		$group += @(,("StoreFront version: ", $stfVersion))

		Output-Table "Server Group Information" $group "list" "Heading2"

		# Get-STFIISSitesInfo cmdlet not available in earlier versions
		try {
			# certificate info
			$https = (Get-STFIISSitesInfo).Binding | ? Protocol -eq "https"
			if ($https -eq $Null) 
				{ Output-Line "*** Missing SSL certificate! ***" }
			else
			{
				$daysToExpire = (New-TimeSpan -End $https.Certificate.NotAfter).Days
				if ($daysToExpire -lt 60) { Output-Line "*** Certificate expires in $($daysToExpire) days!" }
			}
		} catch { }

		# multi-server deployment?
		if ($memberCount -gt 1)
		{
			$servers = @()
			$servers += (,("Server Name", "Status"))
			# get the replication status of each member
			$localMachine = $env:computername
			foreach ($member in $members) 
			{
				if ($member -eq $localMachine)
				{
					$serv = "$($member) (this server)"
					$repStatus = Get-ReplicationStatus $member
				} 
				else
				{
					$serv = $member
					$repStatus = Get-ReplicationStatus $member
				}
				$servers += (,( $serv, $repStatus ))
			}
			Output-Table "Server details" $servers "table" "Heading2"
		}
	}	

	PageBreak
	
	# get all the defined NetScaler instances (used later)
	$gateways = @(Get-STFRoamingGateway)

   Section -Style Heading1 'Stores' {

		# get the stores
		$storeSummary = @()
		$storeSummary += @(,("Name", "Authenticated", "Subscriptions Enabled", "Access"))
		##$stores = @(Get-STFStoreService)
		foreach($store in $stores)
		{
			$name = $store.Name
			$authenticated = Translate-YesNo(!$store.Service.Anonymous)
			$subscriptions = Translate-YesNo(!$store.Service.LockedDown)
			$storeGateways = @($store.Gateways)
			if ($storeGateways.Count -eq 0)
				{ $access = "Internal network only" }
			else
				{ $access = "Internal and external networks" }
			$storeSummary += @(,($name, $authenticated, $subscriptions, $access))
		}
		if ($storeSummary.Count -eq 1) { 
			# only heading rec - no stores
			Output-Line "No stores defined."
		} else {
			Output-Table "Store Summary" $storeSummary "table" "Heading2"
		}

		# individual store details
		foreach($store in $stores)
		{
			# filters, TreatDesktopsAsApps
			$enumOptions = Get-STFStoreEnumerationOptions $store
			# timeouts
			$config = Get-STFStoreFarmConfiguration $store
			$enableFTA = Translate-YesNo($config.EnableFileTypeAssociation)
			$socketPooling = Translate-YesNo($config.PooledSockets)
			$communicationTimeout = $config.CommunicationTimeout
			$connectionTimeout = $config.ConnectionTimeout
			$multiFarmAuth = $config.MultiFarmAuthenticationMode
			$advHealthCheck = Translate-YesNo($config.AdvancedHealthCheck)
			$communicationAttempts = $config.ServerCommunicationAttempts

			$launchOptions = Get-STFStoreLaunchOptions $store
			$farms = @(Get-STFStoreFarm $store)
			$storeDetails = @()
			$url = $baseurl+$store.VirtualPath.Substring(1)
			$PNAURL = $baseURL+$store.VirtualPath.Substring(1)+"/PNAgent/config.xml"

			# roaming account options
			$roaming = Get-STFRoamingAccount -StoreService $store
			$remoteAccessType = Translate-RemoteAccess($roaming.RemoteAccessType)
			$advertised = Translate-YesNo($roaming.Published)
			$authService = Get-STFAuthenticationService $store.AuthenticationServiceVirtualPath
			$tokenValidation = $authService.TokenIssuerUrl
			if ($store.Routing.ExternalEndpoints.Count -eq 0)
				{ $unified = "Disabled" }
			else
				{ $unified = "Enabled" }
			$storeDetails += (,("Store URL:", $url))
			$storeDetails += (,("XenApp Services URL:", $PNAURL))
			$storeDetails += (,("Remote Access:", $remoteAccessType ))
			$storeDetails += (,("Advertised:", $advertised ))
			$storeDetails += (,("Unified Experience:", $unified ))
			$authMethods = 0
			$strMethods = ""
			foreach($authProtocol in $authService.Authentication.ProtocolChoices)
			{
				$strMethod = Translate-Method($authProtocol.Name)
				if ($strMethod -ne "" -and $authProtocol.Enabled)
				{
					++$authMethods            
					if ($authMethods -eq 1)
						{ $strMethods = $strMethod }
					else
						{ $strMethods += ",`r`n$($strMethod)" }
				}
			}
			if ($authMethods -gt 0)
			{
				$storeDetails += (,("Authentication Methods:", $strMethods ))
			}
			
			$storeDetails += (,("Token validation service:", $tokenValidation ))

			Output-Table "$($store.Name): Store Details" $storeDetails "list" "Heading2"

			Output-Line ""

			# process delivery controllers
			$DCSummary = @()
			$DCSummary += (,("Name", "Type", "Protocol", "LB?", "Servers" ))
			foreach ($farm in $farms)
			{
				$farmName = $farm.FarmName
				$farmType = $farm.FarmType
				$transportType = "$($farm.TransportType)"
				if ($transportType -eq "HTTP" -and $farm.Port -ne 80)
					{ $transportType += (" ("+ $farm.Port + ")") }
				elseif ($transportType -eq "HTTPS" -and $farm.Port -ne 443)
					{ $transportType += ("("+ $farm.Port + ")") }
				$LB = Translate-YesNo($farm.LoadBalance)
				$farmServers = @($farm.Servers) -join ","
				$DCSummary += (,($farmName, $farmType, $transportType, $LB, $farmServers ))
			}
			Output-Table "$($store.name): Delivery Controllers" $DCSummary "table" "Heading3"
			Output-Line ""

			# user mapping and site aggregation
			$ufms = @(get-stfuserfarmmapping -storeservice $store)
			if ($ufms.Count -gt 0)
			{
			    $DCAgg = @()
			    $DCAgg += (,("Controller", "Aggregated", "Mapped" ))
			    foreach($farm in $farms)
			    {
			        $mapped = "No"
			        $aggregated = "No"
			        foreach ($farmSet in $ufms.FarmSets)
			        {
			            if ($farmSet.PrimaryFarms -contains $farm.FarmName) 
			            {
			                $mapped = "Yes"
			                if ($farmSet.AggregationGroupName -ne "")
			                    {$aggregated = "Yes"}
			            }
			        }
				$DCAgg += (,($farm.FarmName, $aggregated, $mapped ))
			    }
			    Output-Table "$($store.name): Controller Aggregation" $DCAgg "table" "Heading3"
			    Output-Line ""

			    $aggOptions = @()
			    $aggOptions += (,("LB mode", $ufms[0].Farmsets[0].LoadbalanceMode ))
                if ($ufms[0].Farmsets[0].FarmsAreIdentical -ne $Null) {	
			        $identical = Translate-YesNo($ufms[0].Farmsets[0].FarmsAreIdentical)		
			        $aggOptions += (,("Farms are identical?", $identical ))		
                }	
			    Output-Table "$($store.name): Aggregation Options" $aggOptions "list" "Heading3"
			    Output-Line ""

			    $ufmTable = @()
			    $ufmTable += (,("User Groups", "Primary Farms", "Aggregation Group", "Backup Farms" ))
			    foreach ($ufm in $ufms) 
			    {
				    $userGroups  = $ufm.GroupMembers.Keys -join ",`r`n"
				    $primaryFarms = $ufm.FarmSets.PrimaryFarms -join "`r`n"
				    if ($ufm.FarmSets.AggregationGroupName -eq $Null)
					    { $aggName = " " }
				    else
					    {
                            foreach ($fs in $ufm.FarmSets)
                            {
                                if ($fs.AggregationGroupName -eq 'DefaultAggregationGroup') { $fs.AggregationGroupName = 'Default' }
                                if ($fs.AggregationGroupName -eq '') { $fs.AggregationGroupName = '(not aggregated)' }
                            } 
                            $aggName = ($ufm.FarmSets).AggregationGroupName -join "`r`n " 
                        }
				    if ($ufm.FarmSets.BackupFarms -eq $Null)
					    { $backupFarms = " " }
				    else 
					    { $backupFarms = $ufm.FarmSets.BackupFarms -join ",`r`n" }
				    $ufmTable += (,($userGroups, $primaryFarms, $aggName, $backupFarms  ))
			    }
			    Output-Table "$($store.name): User Farm Mappings" $ufmTable "table" "Heading3"
			    Output-Line ""
			    
			}

			# domains and passwords
			$domainOptions = @()
			$explicitOptions = Get-STFExplicitCommonOptions -AuthenticationService $authService
			$domains = @($explicitOptions.DomainSelection)
			if ($domains.Count -eq 0)
				{ $domainOptions += (,("Users may log on from:", "Any domain" )) }
			else
				{ 
					$domainOptions += (,("Users may log on from:", "Trusted domains only" )) 
					$domainCount = 0
                    $strDomains = ""
					foreach ($domain in $domains)
					{
						++$domainCount
						if ($domainCount -eq 1)
							{ $strDomains = $domain }
						else
							{ $strDomains += "`r`n$($domain)" }
					}
                    if ($domainCount -gt 0)
                    {
                        $domainOptions += (,("Trusted domains:", $strDomains ))
                    }
				}
			Output-Table "$($store.name): Domains" $domainOptions "list" "Heading3"
			Output-Line ""

			# password change options
			$pwOptions = @()

			$pwChange = Translate-PasswordOptions($explicitOptions.AllowUserPasswordChange)
			$pwOptions += (,("Allow password change?", $pwChange ))
			if ($explicitOptions.AllowUserPasswordChange -ne "Never")
			{
				$expiryReminder = Translate-PWReminder($explicitOptions.ShowPasswordExpiryWarning)
				$pwOptions += (,("Password expiry reminder:", $expiryReminder ))
			}
			if ($explicitOptions.Authenticator -eq "defaultDelegatedAuthenticator")
				{ $authenticator = "Active Directory" }
			else
				{ $authenticator = "Delivery Controllers" }
			$pwOptions += (,("Validate passwords via:", $authenticator ))
			Output-Table "$($store.name): Password options" $pwOptions "list" "Heading3"
			Output-Line ""

			#optimal gateway routing
			$optimalRouting = @(Get-STFStoreRegisteredOptimalLaunchGateway $store)
			if ($optimalRouting.Count -gt 0)
			{
				$gatewayRouting = @()
				$gatewayRouting += (,("Optimal gateway", "Direct access", "Farms", "Zones" ))
				foreach($route in $optimalRouting)
				{
					$optGateway = $route.Name
					if ($optGateway -eq "_") { $optGateway = "Direct HDX connection" }
					$directAccess = Translate-YesNo($route.EnabledOnDirectAccess)
					if (@($route.Farms).Count -eq 0) 
						{ $optGatewayFarms = " " }
					else 	
						{ $optGatewayFarms = @($route.Farms) -join "," }
					if (@($route.Zones).Count -eq 0) 
						{ $zones = " " }
					else
						{ $zones = @($route.Zones) -join "," }

					$gatewayRouting += (,($optGateway, $directAccess, $optGatewayFarms, $zones ))
				}
				Output-Table "$($store.name): Optimal gateway routing" $gatewayRouting "table" "Heading3"
			}
            
            # store enumeration options
            $enumOptions = @()
			$subtituteDesktopImage = Translate-YesNo($store.Service.SubstituteDesktopImage)
			$enhancedEnum = Translate-YesNo($store.Resources.Enumeration.EnhancedEnumeration)
			$desktopsAsApps = Translate-YesNo($store.Resources.Enumeration.TreatDesktopsAsApps)
			$filterByTypes = $store.Resources.Enumeration.FilterByTypesInclude
			$filterByKeyInclude = $store.Resources.Enumeration.FilterByKeywordsInclude
			$filterByKeyExclude = $store.Resources.Enumeration.FilterByKeywordsExclude
			$enumOptions += (,("Substitute desktop image:", $subtituteDesktopImage ))
			$enumOptions += (,("Enable enhanced enumeration:", $enhancedEnum ))
			$enumOptions += (,("Treat desktops as apps:", $desktopsAsApps ))
			$enumOptions += (,("Filter resources by type:", $filterByTypes ))
			$enumOptions += (,("Filter resources by excluded keywords:", $filterByKeyExclude ))
			$enumOptions += (,("Filter resources by included keywords:", $filterByKeyInclude ))
			$enumOptions += (,("Socket pooling:", $socketPooling ))
			$enumOptions += (,("Communication timeout:", $communicationTimeout ))
			$enumOptions += (,("Connection timeout:", $connectionTimeout ))
			$enumOptions += (,("Advanced health check:", $advHealthCheck ))
			$enumOptions += (,("Server communication attempts:", $communicationAttempts ))
			$enumOptions += (,("Multi-farm authentication mode:", $multiFarmAuth ))
 
			Output-Table "$($store.name): Enumeration Options" $enumOptions "list" "Heading3"
			Output-Line ""

			# launch options
			$launchOptions = @()
			$allowSessionReconnect = Translate-YesNo($store.Service.AllowSessionReconnect)
			$launchOptions += (,("Allow session reconnect:", $allowSessionReconnect ))

			$addressResolution = $store.Resources.Launch.AddressResolutionType
			$overrideClientName = Translate-YesNo($store.Resources.Launch.OverrideIcaClientName)
			$showDesktopViewer = Translate-YesNo($store.Resources.Launch.ShowDesktopViewer)
			$allowFolderRedirect = Translate-YesNo($store.Resources.Launch.AllowSpecialFolderRedirection)
			$requireLaunchRef = Translate-YesNo($store.Resources.Launch.RequireLaunchReference)
			$launchOptions += (,("Address resolution type:", $addressResolution ))
			$launchOptions += (,("Override ICA client name:", $overrideClientName ))
			$launchOptions += (,("Show desktop viewer:", $showDesktopViewer ))
			$launchOptions += (,("Allow special folder redirection:", $allowFolderRedirect ))
			$launchOptions += (,("Require launch reference:", $requireLaunchRef ))
			$launchOptions += (,("Enable File Type Association:", $enableFTA ))
 
			Output-Table "$($store.name): Launch Options" $launchOptions "list" "Heading3"
			Output-Line ""

            PageBreak;

			# Receiver for Web site summary
			$siteSummary = @()
			$siteSummary += (,("Receiver for Web Site", "Classic Experience", "Authentication Methods", "HTML5" ))
			$sites = @(Get-STFWebreceiverService -StoreService $store)

			foreach ($site in $sites)
			{
				$siteName = $baseURL+$site.VirtualPath.Substring(1)
				$classic = Translate-Enabled($site.WebReceiver.ClassicReceiverExperience)
				$authMethods = @((Get-STFWebReceiverAuthenticationMethods -WebReceiverService $site).Methods)
				$receiver = Get-STFWebReceiverPluginAssistant -WebReceiverService $site
				$HTML5ver = $receiver.Html5.Version
				if ($HTML5ver -eq "0.0.0.0" -or $HTML5ver -eq $Null) { $HTML5ver = "Not Used" }
                $strMethods = ""
				for ($idx=0; $idx -lt $authMethods.Count; $idx++)
				{ if ($idx -eq 0) { $strMethods = Translate-Method($authMethods[$idx]) } 
					else {
						$strMethods += ",`r`n" 
						$strMethods += Translate-Method($authMethods[$idx]) 
					} 
				}
				$siteSummary += (,($siteName, $classic, $strMethods, $HTML5ver ))
			}
			if ($siteSummary.Count -eq 1) {
				Output-Line "No web sites defined."
			} else {
				Output-Table "$($store.name): Web Site Summary" $siteSummary "table" "Heading2"
			}
			Output-Line ""

			# Receiver for Web site details
			foreach ($site in $sites)
			{
				# application groups
				$appGroups = @(Get-STFWebReceiverFeaturedAppGroup $site)
				if ($appGroups.Count -gt 0)
				{
					$appGroupInfo = @()
					$appGroupInfo += (,("Title", "Description", "Image ID", "Type", "Contents" ))
					foreach($appGroup in $appGroups)
					{
						$title = $appGroup.Title
						$desc = $appGroup.Description
						$imageID = $appGroup.TileId 
						$type = $appGroup.ContentType
						$contents = @($appGroup.Contents)
                        $strContents = ""
 				        if ($contents.Count -gt 0) 
				        {
                            for ($idx=0; $idx -lt $contents.Count; $idx++)
                            {
                                if ($idx -eq 0) {$strContents = $contents[$idx]}
                                else {$strContents += "`r`n$($contents[$idx])"}
                            }
                        }                       
						$appGroupInfo += (,($title, $desc, $imageID, $type, $strContents ))
					}
					Output-Table "$($store.name) : $($site.Name) - Application Groups" $appGroupInfo "table" "Heading3"
					Output-Line ""

					# background images for the application groups
					$tiles = @(Get-STFWebReceiverFeaturedAppGroupTiles $site)
					$tileInfo = @()
					$tileInfo += (,("Image ID", "Background Image" ))
					foreach($tile in $tiles)
					{
						$tileID = $tile.TileId
						$image = $tile.BackgroundImage
						$tileInfo += (,($tileID, $image ))
					}
					Output-Table "$($store.name) : $($site.Name) - Application Group Tiles" $tileInfo "table" "Heading3"
					Output-Line ""
				}
			
				$ciDetails = @()               # client interface
				$deployDetails = @()           # deployment  
				$wcDetails = @()               # workspace control 
				$miscDetails = @()             # misc options 
    
				$style = Get-STFWebReceiverSiteStyle -WebReceiverService $site
				$classic = Translate-Enabled($site.WebReceiver.ClassicReceiverExperience)
				$ciDetails += (,("Classic experience:", $classic ))
				if ($site.WebReceiver.ClassicReceiverExperience -eq $False)
				{
					$logonLogo = $style.HeaderLogoPath
					$appLogo = $style.HeaderLogoPath
					$linkColor = $style.LinkColor
					$headerFG = $style.HeaderForegroundColor
					$headerBG = $style.HeaderBackgroundColor
					$ciDetails += (,("Logon page logo:", $logonLogo ))
					$ciDetails += (,("Application page logo:", $appLogo ))
					$ciDetails += (,("Header BG color:", $headerBG ))
					$ciDetails += (,("Header FG color:", $headerFG ))
					$ciDetails += (,("Link color:", $linkColor ))
				}

				$authMethods = @((Get-STFWebReceiverAuthenticationMethods -WebReceiverService $site).Methods)
				if ($authMethods.Count -gt 0) 
				{
					for($idx=0; $idx -lt $authMethods.Count; $idx++)
					{
						
						if ($idx -eq 0) { $strMethods = Translate-Method($authMethods[$idx]) }
						else { 
                            $strMethods += ",`r`n"
                            $strMethods += Translate-Method($authMethods[$idx]) 
                        }
					}
                    $miscDetails += (,("Authentication methods:", $strMethods ))
				}

				$appShortcuts = @((Get-STFWebReceiverApplicationShortcuts -WebReceiverService $site).TrustedUrls)
				if ($appShortcuts.Count -gt 0) 
				{
                    $URLs = ""
					for($idx=0; $idx -lt $appShortcuts.Count; $idx++)
					{
						if ($idx -eq 0) { $URLs = $appShortcuts[$idx].OriginalString }
						else { $URLs += "`r`n$($appShortcuts[$idx].OriginalString)" }
					}
                    $ciDetails += (,("Shortcut URLs:", $URLs ))
				}

				#Citrix Receiver
				$receiver = Get-STFWebReceiverPluginAssistant -WebReceiverService $site
				$deployment = Translate-HTML5Deployment($receiver.Html5.Enabled)
                if ($deployment -ne $Null) {
				    $deployDetails += (,("Receiver deployment:", $deployment ))
                }
				if ($receiver.Html5.Enabled -ne "Off")
				{
                    if ($receiver.Html5.Version -ne $Null) {
					    $deployDetails += (,("HTML5 version:", $receiver.Html5.Version ))
                    }
					$sameWindow = Translate-YesNo($receiver.Html5.SingleTabLaunch)
                    if ($sameWindow -ne $Null) {
					    $deployDetails += (,("Launch HTML5 in same window:", $sameWindow ))
                    }
				}
				$downloadPlugin = Translate-YesNo($receiver.Enabled)
                if ($downloadPlugin -ne $Null) {
				    $deployDetails += (,("Allow Receiver download:", $downloadPlugin ))
                }
				if ($receiver.Enabled)
				{
					$upgradeAtLogin = Translate-YesNo($receiver.UpgradeAtLogin)
                    if ($upgradeAtLogin -ne $Null) {
					    $deployDetails += (,("Upgrade Receiver at login:", $upgradeAtLogin ))
                    }
					$WinSource = $receiver.Win32.Path
					$MacSource = $receiver.MacOS.Path
					if ($WinSource.Contains("downloadplugins.citrix.com")) { $WinSource = "Citrix Web Site" }
					if ($MacSource.Contains("downloadplugins.citrix.com")) { $MacSource = "Citrix Web Site" }
                    if ($WinSource -ne $Null) {
					    $deployDetails += (,("Windows Receiver source:", $WinSource ))
                    }
                    if ($MacSource -ne $Null) {
					    $deployDetails += (,("MacOS Receiver source:", $MacSource ))
                    }
				}

				$webResource = Get-STFWebReceiverResourcesService -WebReceiverService $site
				$showToolBar = Translate-YesNo($webResource.ShowDesktopViewer)

				$clientInterface = Get-STFWebReceiverUserInterface -WebReceiverService $site

				# workspace control
				$logoffAction = $clientInterface.WorkspaceControl.LogoffAction
				$enableWC = Translate-YesNo($clientInterface.WorkspaceControl.Enabled)
				$reconnectAtLogon = Translate-YesNo($clientInterface.WorkspaceControl.AutoReconnectAtLogon)
				$showReconnect = Translate-YesNo($clientInterface.WorkspaceControl.ShowReconnectButton)
				$showDisconnect = Translate-YesNo($clientInterface.WorkspaceControl.ShowDisconnectButton)
				$wcDetails += (,("Logoff action:", $logoffAction ))
				$wcDetails += (,("Workspace control enabled:", $enableWC ))
                if ($enableWC -eq "Yes")
                {
				    $wcDetails += (,("Reconnect sessions at logon:", $reconnectAtLogon ))
				    $wcDetails += (,("Show reconnect button:", $showReconnect ))
				    $wcDetails += (,("Show disconnect button:", $showDisconnect ))
                }

				# client interface settings
				$autoLaunch = Translate-YesNo($clientInterface.AutoLaunchDesktop)
				$receiverConfig = Translate-YesNo($clientInterface.ReceiverConfiguration.Enabled)
				$multiClick = "$($clientInterface.MultiClickTimeout) seconds"
				$showApps = Translate-YesNo($clientInterface.UIViews.ShowAppsView)
				$showDesktops = Translate-YesNo($clientInterface.UIViews.ShowDesktopsView)
				$defaultView = $clientInterface.UIViews.DefaultView
				$ciDetails += (,("Auto launch desktop:", $autoLaunch ))
				$ciDetails += (,("ShowDesktopViewer:", $showToolBar ))
				$ciDetails += (,("Enable Receiver configuration:", $receiverConfig ))
				$ciDetails += (,("Multi-click duration:", $multiClick ))
				$ciDetails += (,("Show Apps view:", $showApps ))
				$ciDetails += (,("Show Desktops view:", $showDesktops ))
				$ciDetails += (,("Default view:", $defaultView ))

				# misc web site info
				$sessionTimeout = $site.WebReceiver.SessionStateTimeout
                if ($sessionTimeout -ne $Null) {
				    $miscDetails += (,("Session timeout:", $sessionTimeout ))
                }
				$communication = Get-STFWebReceiverCommunication -WebReceiverService $site
				$authManager = @(Get-STFWebReceiverAuthenticationManager -WebReceiverService $site)
				if ($authManager[0].LoginFormTimeout -ne $Null)
				{
					$authTimeout = $authManager[0].LoginFormTimeout
					$miscDetails += (,("Logon form timeout (mins.):", $authTimeout )) 
				}
				$iconExpiry = $webResource.IcaFileCacheExpiry
				$iconSize = $webResource.IconSize
				$miscDetails += (,("Icon Cache Expiry:", $iconExpiry ))
				$ciDetails += (,("Icon Resolution:", $iconSize ))

				Output-Table "$($store.name) : $($site.Name) - Client Interface" $ciDetails "list" "Heading3"
				Output-Line ""
				Output-Table "$($store.name) : $($site.Name) - Deployment" $deployDetails "list" "Heading3"
				Output-Line ""
				Output-Table "$($store.name) : $($site.Name) - Workspace Control" $wcDetails "list" "Heading3"
				Output-Line ""
				Output-Table "$($store.name) : $($site.Name) - Misc Options" $miscDetails "list" "Heading3"
				Output-Line ""
			}

			# Gateways (remote access settings) for the store
			$storeGateways = @($store.Gateways)
			$gatewaySummary = @()
			$gatewaySummary += (,("Display Name", "Version", "URL" ))
			foreach ($storeGateway in $storeGateways)
			{
				# look up the store's gateway in the full gateway table 
				$gw = $gateways | ? Id -eq $storeGateway.Key
				$gwName = $gw.Name
				$gwVersion = Translate-NSVersion($gw.Version)
				$gwURL = $gw.Location
				$gatewaySummary += (,($gwName, $gwVersion, $gwURL ))
			}
			if ($gatewaySummary.Count -eq 1) {
				Output-Line "$($store.name) : No gateways registered for this store."
			} else {
				Output-Table "$($store.name) : $($site.Name) - NetScaler Instances" $gatewaySummary "table" "Heading3"
			}
			Output-Line ""

			# XenApp Services Support
			$pna = Get-STFStorePna $store
			$pnaSupport = @()
			$pnaEnabled = Translate-YesNo($pna.PnaEnabled)
			$pnaDefault = Translate-YesNo($pna.DefaultPnaService)
			$pnaURL = "$($baseURL)Citrix/$($store.Name)/PNAgent/config.xml"
			$pnaSupport += (,("XenApp Services support enabled:", $pnaEnabled ))
			$pnaSupport += (,("Default store for XenApp Services:", $pnaDefault ))
			$pnaSupport += (,("XenApp Services URL:", $PNAURL ))
			Output-Table "$($store.name) - XenApp Services Support" $pnaSupport "list" "Heading3"
			Output-Line ""

			}
		}	### end of store information

       PageBreak;
		
	   Section -Style Heading1 'NetScaler Gateways' {

        ##$stores = @(Get-STFStoreService)
		# ALL NetScaler instances
		$gatewaySummary = @()
		$gatewaySummary += (,("Display Name", "Role", "Used by", "URL", "GSLB URL" ))
		foreach ($gw in $gateways)
		{
			$gwName = $gw.Name
			$gwRole = "Auth and HDX routing"
			if ($gw.AuthenticationCapable -eq $False)
				{ $gwRole = "HDX routing only" }
			elseif ($gw.HdxRoutingCapable -eq $False)
				{ $gwRole = "Authentication only" }
			$gwURL = $gw.Location
			$gslbURL = $gw.GslbLocation
			if ($gslbURL -eq $Null) {$gslbURL = " "}
			# check which stores use this gateway
			$gwStores = ""
			foreach($store in $stores)
			{
				if ($store.Gateways.Key -contains $gw.Id)
                {
                    if ($gwStores -eq "")
				        { $gwStores = $store.Name }
                    else
                        { $gwStores += "`r`n$($store.Name)" }
                }
			}
			$gatewaySummary += (,($gwName, $gwRole, $gwStores, $gwURL, $gslbURL ))
		}
		if ($gatewaySummary.Count -eq 1) {
			Output-Line "No gateways defined."
		} else {
			Output-Table "NetScaler Instances - Summary" $gatewaySummary "table" "Heading1"
		}
		Output-Line ""

		# NetScaler instance details
		foreach ($gw in $gateways)
		{
			$gwDetails = @()
			$gwName = $gw.Name
			$gwRole = "Authentication and HDX routing"
			if ($gw.AuthenticationCapable -eq $False)
				{ $gwRole = "HDX routing only" }
			elseif ($gw.HdxRoutingCapable -eq $False)
				{ $gwRole = "Authentication only" }
			$gwURL = $gw.Location
			$gslbURL = $gw.GslbLocation
			if ($gslbURL -eq $Null) {$gslbURL = " "}
			$gwDetails += (,("ID:", $gw.Id ))
			$gwDetails += (,("Role:", $gwRole ))
			$gwDetails += (,("URL:", $gwURL ))
			$gwDetails += (,("GSLB URL:", $gslbURL ))

			$stas = @($gw.SecureTicketAuthorityUrls)
			$staCount = 0
			foreach ($sta in $stas)
			{
				++$staCount
                # older version of StoreFront may not have .Url
                if ($sta.Url -eq $Null) {
                    $staURL = $sta.ToString()
                } else {
				    $staURL = $sta.Url.ToString()
                }
				if ($staCount -eq 1)
					{ $staURLs = $staURL }
				else
					{ $staURLs += "`r`n$($staURL)" }
			}
            if ($staCount -gt 0)
                { $gwDetails += (,("STA URLs:", $staURLs )) }
			$lbSTAs = Translate-YesNo($gw.StasUseLoadBalancing)
			$srSTAs = Translate-YesNo($gw.SessionReliability)
			$bypassDuration = $gw.StasBypassDuration
			$twoTickets = Translate-YesNo($gw.RequestTicketTwoStas)
			$NSVersion = Translate-NSVersion($gw.Version)
			$logonType = Translate-LogonType($gw.Logon)
			$NSVIP = "$($gw.IpAddress)"
			$fallback = $gw.SmartCardFallback
			$callbackURL = "$($gw.CallbackUrl)"
			$gwDetails += (,("Load balance multiple STA servers:", $lbSTAs ))
			$gwDetails += (,("Enable session reliability:", $srSTAs ))
			$gwDetails += (,("Bypass failed STA for:", $bypassDuration  ))
			$gwDetails += (,("Request tickets from two STAs:", $twoTickets ))
			$gwDetails += (,("NetScaler version:", $NSVersion ))
			$gwDetails += (,("vServer IP:", $NSVIP ))
			$gwDetails += (,("Logon type:", $logonType ))
			$gwDetails += (,("Smart card fallback:", $fallback ))
			$gwDetails += (,("Callback URL:", $callbackURL ))
			Output-Table "$($gwName)" $gwDetails "list" "Heading2"
			Output-Line ""
		}
	}	### end of gateway information
	
   PageBreak

   Section -Style Heading1 'Beacon Information' {
	
		# beacons
		$beaconList = @()
		$beacons = Get-STFRoamingBeacon -Internal
		$beaconList += (,("Internal beacon:", $beacons ))
		$beacons = @(Get-STFRoamingBeacon -External)
		$beaconCnt = 0
		foreach($beacon in $beacons)
		{
			++$beaconCnt
			if ($beaconCnt -eq 1)
				{ $strBeacons = $beacon }
			else
				{ $strBeacons += "`r`n$($beacon)" }
		}
        if ($beaconCnt -gt 0)
            { $beaconList += (,("External beacons:", $strBeacons )) }

	Output-Table "Beacons" $beaconList "list" "Heading2"
   }	### end of beacon information
	
    # process hardware and software info for the server group

    $memberServers = @($serverGroup.ClusterMembers).Server

    if ($Script:Hardware -eq $True) {
		foreach ($member in $memberServers) 
		{
	        PageBreak
            Section -Style Heading1 "$($member) - Hardware" {
                GetComputerWMIInfo $member
            }
        }
    }

    if ($Script:Software -eq $True) {
		foreach ($member in $memberServers) 
		{
	        PageBreak
            Section -Style Heading1 "$($member) - Software" {
                InstalledSoftware $member
	            CitrixServices $member
            }
        }
    }

}

<#  Generate Pscribo output #>
$outputFormats = @()
if ($Script:OutputWord -eq $True) { $outputFormats += "Word" }
if ($Script:OutputHTML -eq $True) { $outputFormats += "HTML" }
if ($Script:OutputText -eq $True) { $outputFormats += "Text" }
$document | Export-Document -Path $Script:OutputDir -Format $outputFormats ;
Write-Host "$(Get-Date): Done!"
#endregion SFDoc
#


# SIG # Begin signature block
# MIIcaAYJKoZIhvcNAQcCoIIcWTCCHFUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVUBbWjlr7oc9Q/mWo5iXIxMz
# T1+ggheXMIIFIDCCBAigAwIBAgIQB1tuZV4A7CBDAwTuuFjowzANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MTAwODAwMDAwMFoXDTIwMTAx
# NTEyMDAwMFowXTELMAkGA1UEBhMCVVMxETAPBgNVBAgTCE5ldyBZb3JrMREwDwYD
# VQQHEwhCcm9va2x5bjETMBEGA1UEChMKU2FtIEphY29iczETMBEGA1UEAxMKU2Ft
# IEphY29iczCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMnDYUmYWwbP
# T7mW7CAerbErENc2DwfYP8FPEYcyEmR9uEd7Dedd7wM+m2Ikkluadyz/Ui8N5Dob
# tkTJjQjCLdF7sZnPsQKsS9krm8Ml5y00nnNOX/E7aE04tHr0cojAMQlWCQiN8Wcz
# QW+m/6s4kfVwGfYZCZAqPs5flKFVjx8DWvBO5FJ8V1cXMQY8WNPW/oeWONJtOKLN
# hSzMoP7y0EFBaE+iGEnVVKHEzhEWIREjak4vdRGpQHvrQx8yciC3Sf9BJ1MEsaGt
# 9wcwo1j5R5gS9Q7xC+lqHXJ06mauJqEUvpK+a+fO17pZ9jnhWK9PDGnDX01eRgG7
# fsY53TiAZVkCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5
# LfZldQ5YMB0GA1UdDgQWBBSJZ2ZtFAqJVxC2QvxoS9veoDry3TAOBgNVHQ8BAf8E
# BAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAz
# oDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEu
# Y3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBz
# Oi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4
# MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEF
# BQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFz
# c3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcN
# AQELBQADggEBAKka7JsKvYTkQoBVbP4P6Mpw00ujba8oF0xtxv2qaug9+GbrIX4c
# 1cRLm3AD9fozDbUGXJ3M1WqVWU2tM3x6jaWtV0GxcueSBWLeP4trIGItzvdQggWa
# NdfCvon1JoshzLB1t/ao8UTvHd/3zRXx2TmpnCqN69CfgiWeGPo7p0m4cwDTVqSR
# sxcD4V/IcDSfq6B1Ob5s087g9qpBrGeAs3vDidi8DsM4aj0Cgnxu9P7KK71aSTr/
# Rw4UdEH845+rGxmgVJ3eIFR3fYMDo40y0rrWckrWA+I37p5NZEdh4mRqZRGAGM0Q
# TCi9wgr1gWZNq9rr7J2rfhwzRd3gboaDkaEwggUwMIIEGKADAgECAhAECRgbX9W7
# ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBa
# Fw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/l
# qJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fT
# eyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqH
# CN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+
# bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLo
# LFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIB
# yTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAK
# BggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwA
# AgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAK
# BghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0j
# BBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7s
# DVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGS
# dQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6
# r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo
# +MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qz
# sIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHq
# aGxEMrJmoecYpJpkUe8wggZqMIIFUqADAgECAhADAZoCOv9YsWvW1ermF/BmMA0G
# CSqGSIb3DQEBBQUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IEFzc3VyZWQgSUQgQ0EtMTAeFw0xNDEwMjIwMDAwMDBaFw0yNDEwMjIwMDAwMDBa
# MEcxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhEaWdpQ2VydDElMCMGA1UEAxMcRGln
# aUNlcnQgVGltZXN0YW1wIFJlc3BvbmRlcjCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAKNkXfx8s+CCNeDg9sYq5kl1O8xu4FOpnx9kWeZ8a39rjJ1V+JLj
# ntVaY1sCSVDZg85vZu7dy4XpX6X51Id0iEQ7Gcnl9ZGfxhQ5rCTqqEsskYnMXij0
# ZLZQt/USs3OWCmejvmGfrvP9Enh1DqZbFP1FI46GRFV9GIYFjFWHeUhG98oOjafe
# Tl/iqLYtWQJhiGFyGGi5uHzu5uc0LzF3gTAfuzYBje8n4/ea8EwxZI3j6/oZh6h+
# z+yMDDZbesF6uHjHyQYuRhDIjegEYNu8c3T6Ttj+qkDxss5wRoPp2kChWTrZFQlX
# mVYwk/PJYczQCMxr7GJCkawCwO+k8IkRj3cCAwEAAaOCAzUwggMxMA4GA1UdDwEB
# /wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMIIB
# vwYDVR0gBIIBtjCCAbIwggGhBglghkgBhv1sBwEwggGSMCgGCCsGAQUFBwIBFhxo
# dHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMIIBZAYIKwYBBQUHAgIwggFWHoIB
# UgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkA
# YwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEA
# bgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMA
# UABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkA
# IABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwA
# aQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8A
# cgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMA
# ZQAuMAsGCWCGSAGG/WwDFTAfBgNVHSMEGDAWgBQVABIrE5iymQftHt+ivlcNK2cC
# zTAdBgNVHQ4EFgQUYVpNJLZJMp1KKnkag0v0HonByn0wfQYDVR0fBHYwdDA4oDag
# NIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0Et
# MS5jcmwwOKA2oDSGMmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRENBLTEuY3JsMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNydDANBgkqhkiG
# 9w0BAQUFAAOCAQEAnSV+GzNNsiaBXJuGziMgD4CH5Yj//7HUaiwx7ToXGXEXzakb
# vFoWOQCd42yE5FpA+94GAYw3+puxnSR+/iCkV61bt5qwYCbqaVchXTQvH3Gwg5QZ
# BWs1kBCge5fH9j/n4hFBpr1i2fAnPTgdKG86Ugnw7HBi02JLsOBzppLA044x2C/j
# bRcTBu7kA7YUq/OPQ6dxnSHdFMoVXZJB2vkPgdGZdA0mxA5/G7X1oPHGdwYoFenY
# k+VVFvC7Cqsc21xIJ2bIo4sKHOWV2q7ELlmgYd3a822iYemKC23sEhi991VUQAOS
# K2vCUcIKSK+w1G7g9BQKOhvjjz3Kr2qNe9zYRDCCBs0wggW1oAMCAQICEAb9+QOW
# A63qAArrPye7uhswDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTA2MTExMDAwMDAw
# MFoXDTIxMTExMDAwMDAwMFowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGln
# aUNlcnQgQXNzdXJlZCBJRCBDQS0xMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEA6IItmfnKwkKVpYBzQHDSnlZUXKnE0kEGj8kz/E1FkVyBn+0snPgWWd+e
# tSQVwpi5tHdJ3InECtqvy15r7a2wcTHrzzpADEZNk+yLejYIA6sMNP4YSYL+x8cx
# SIB8HqIPkg5QycaH6zY/2DDD/6b3+6LNb3Mj/qxWBZDwMiEWicZwiPkFl32jx0Pd
# Aug7Pe2xQaPtP77blUjE7h6z8rwMK5nQxl0SQoHhg26Ccz8mSxSQrllmCsSNvtLO
# Bq6thG9IhJtPQLnxTPKvmPv2zkBdXPao8S+v7Iki8msYZbHBc63X8djPHgp0XEK4
# aH631XcKJ1Z8D2KkPzIUYJX9BwSiCQIDAQABo4IDejCCA3YwDgYDVR0PAQH/BAQD
# AgGGMDsGA1UdJQQ0MDIGCCsGAQUFBwMBBggrBgEFBQcDAgYIKwYBBQUHAwMGCCsG
# AQUFBwMEBggrBgEFBQcDCDCCAdIGA1UdIASCAckwggHFMIIBtAYKYIZIAYb9bAAB
# BDCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1j
# cHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAA
# dQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAA
# YwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8A
# ZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4A
# ZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUA
# ZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwA
# aQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQA
# IABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZI
# AYb9bAMVMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0O
# BBYEFBUAEisTmLKZB+0e36K+Vw0rZwLNMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1R
# i6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQBGUD7Jtygkpzgdtlspr1LPUukx
# R6tWXHvVDQtBs+/sdR90OPKyXGGinJXDUOSCuSPRujqGcq04eKx1XRcXNHJHhZRW
# 0eu7NoR3zCSl8wQZVann4+erYs37iy2QwsDStZS9Xk+xBdIOPRqpFFumhjFiqKgz
# 5Js5p8T1zh14dpQlc+Qqq8+cdkvtX8JLFuRLcEwAiR78xXm8TBJX/l/hHrwCXaj+
# +wc4Tw3GXZG5D2dFzdaD7eeSDY2xaYxP+1ngIw/Sqq4AfO6cQg7PkdcntxbuD8O9
# fAqg7iwIVYUiuOsYGk38KiGtSTGDR5V3cdyxG0tLHBCcdxTBnU8vWpUIKRAmMYIE
# OzCCBDcCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IElu
# YzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQg
# U0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQQIQB1tuZV4A7CBDAwTuuFjo
# wzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTAjBgkqhkiG9w0BCQQxFgQU66utw3ytUuSb4ZD4K03J9w0XSigwDQYJKoZIhvcN
# AQEBBQAEggEAQPeSIuVLp5YesjX/NI9wO5MBTG1hxCDEwQinNorDYd0NI9FN1hi2
# dlgyjRbdmIe5EU1D1mzYrMpOR+pJjioGORei5jl3Dv9snDYV7hH9mznO7/ntB9Eu
# y3vYO+Ie3b5YEgjtfz+DzQlxcdWrT84RmbrdeVvBhvtzBPFlMmQtElttdpWd6ttd
# Cilx4NWMlOTFc20etcKH12w2VDyzpSovW+y8ixeQTaczbXp5+nrhrmEHKIxTS0BA
# z/Pa1KgpKFw8eNiSEUMco3sDzSXuAej3IQFkefxE5UC3QgEr/NulLzHNdx4gsNPX
# bDzq1Zn8whCLYffx8oWrURT1ZgLTVsylYaGCAg8wggILBgkqhkiG9w0BCQYxggH8
# MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IEFz
# c3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfwZjAJBgUrDgMCGgUAoF0wGAYJ
# KoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTkxMDIzMTg0
# MTE0WjAjBgkqhkiG9w0BCQQxFgQUn0NH5hIAobTTCjAvhV2UWPBA2e4wDQYJKoZI
# hvcNAQEBBQAEggEAMaFLrG9K4z/q7rKoNnOm97wMJ/EMsN6K4c18j2OoCzkBHYPJ
# JTYEZmuS8kLYw2576SF+PSIqZY54WCLOX8KoFPq0oBnQ1bVcVJJsn3lLOl5RILTQ
# E/Nh9O4PIxOwKz9/7vqdS+7iWPs6d7OZ3XBm9hnIXhTF7hHh0ZmeP+lzfox7wguU
# wLkOeechXyjKpV+ECtEk+2sRXQAwtIiBIcsaSK2ZRAyPGJwrNoHJ8mYco5Hm1Y7a
# CI0EIOnl2w40mKJEjnuLmJ0iGqXvmoTupLs71wFk+/liLrXFQCgZYsSMpGHWUWxw
# Wf6jFVCjZN3jZLYSOQjFqDyOVlstWZde6VWQag==
# SIG # End signature block
