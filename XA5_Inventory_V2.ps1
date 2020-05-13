<#
.SYNOPSIS
	Creates a complete inventory of a Citrix XenApp 5 farm using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix XenApp 5 farm using Microsoft Word and PowerShell.
	Works for XenApp 5 Server 2003 32-bit and 64-bit and XenApp 5 Server 2008 32-bit and 64-bit
	Works for Presentation Server 4.5 Server 2003 32-bit and 64-bit
	Creates a Word document named after the XenApp 5 farm.
	Document includes a Cover Page, Table of Contents and Footer.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2007/2010. Mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Grid (Word 2010.Works)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Sideline (Word 2007/2010. Works)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
	Default value is Conservative.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.EXAMPLE
	PS C:\PSScript > .\XA5_Inventory_v2.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Conservative for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA5_Inventory_v2.ps1 -verbose
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Conservative for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript .\XA5_Inventory_v2.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\XA5_Inventory_v2.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.LINK
	http://www.carlwebster.com/documenting-a-citrix-xenapp-5-farm-with-microsoft-powershell-and-word-version-2
.NOTES
	NAME: XA5_Inventory_V2.ps1
	VERSION: 2.02
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith and Jeff Wouters)
	LASTEDIT: June 10, 2013
#>


#thanks to @jeffwouters for helping me with these parameters
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]

Param([parameter(
	Position = 0, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[string]$CompanyName="",
    
	[parameter(
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Motion", 

	[parameter(
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username )


Set-StrictMode -Version 2

#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mikebogo on Twitter
#The original script was designed to be run on a XenApp 6 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from original script for XenApp 5
#originally released to the Citrix community on October 3, 2011
#updated October 9, 2011.  Added CPU Utilization Management, Memory Optimization and Health Monitoring & Recovery
#updated January 26, 2013 to output to Microsoft Word 2007 and 2010
#	Test for CompanyName in two different registry locations
#	Fixed issues found by running in set-strictmode -version 2.0
#	Fixed typos
#	Test if template DOTX file loads properly.  If not, skip Cover Page and Table of Contents
#	Add more write-verbose statements
#	Disable Spell and Grammer Check to resolve issue and improve performance (from Pat Coughlin)
#	Added in the missing Load evaluator settings for Load Throttling and Server User Load 
#	Test XenApp server for availability before getting services and hotfixes
#	Move table of Citrix services to align with text above table
#	Created a table for Citrix installed hotfixes
#	Created a table for Microsoft hotfixes
#Updated March 14, 2013
#	?{?_.SessionId -eq $SessionID} should have been ?{$_.SessionId -eq $SessionID} in the CheckWordPrereq function
#Updated April 20, 2013
#	Fixed five typos dealing with Session Printer policy settings
#	Fixed a compatibility issue with the way the Word file was saved and Set-StrictMode -Version 2
#Updated June 7, 2013
#	Fixed the content of and the detail contained in the Table of Contents
#	Citrix services that are Stopped will now show in a Red cell with bold, black text
#	Added a few more Write-Verbose statements

Function CheckWordPrereq
{
	if ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $null
	if ($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

Function ValidateCompanyName
{
	$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
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
# This function just gets $true or $false
function Test-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $null -ne $key.GetValue($name, $null)
}

# Gets the specified registry value or $null if it is missing
function Get-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    if ($key) {
        $key.GetValue($name, $null)
    }
}

Function ValidateCoverPage
{
	Param( [int]$xWordVersion, [string]$xCP )
	
	$xArray = ""
	If( $xWordVersion -eq 14)
	{
		#word 2010
		$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
	}
	ElseIf( $xWordVersion -eq 12)
	{
		#word 2007
		$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend" )
	}
	
	If ($xArray -contains $xCP)
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}

Function Check-NeededPSSnapins
{
	Param( [parameter(Mandatory = $true)][alias("Snapin")][string[]]$Snapins)
	
	#function specifics
	$MissingSnapins=@()
	$FoundMissingSnapin=$false
	$loadedSnapins = @()
	$registeredSnapins = @()
    
	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}
    
    foreach ($Snapin in $Snapins){
        #check if the snapin is loaded
        if (!($LoadedSnapins -like $snapin)){

            #Check if the snapin is missing
            if (!($RegisteredSnapins -like $Snapin)){

                #set the flag if it's not already
                if (!($FoundMissingSnapin)){
                    $FoundMissingSnapin = $True
                }
                
                #add the entry to the list
                $MissingSnapins += $Snapin
            }#End Registered If 
            
            Else{
                #Snapin is registered, but not loaded, loading it now:
                Write-Host "Loading Windows PowerShell snap-in: $snapin"
                Add-PSSnapin -Name $snapin
            }
            
        }#End Loaded If
        #Snapin is registered and loaded
        else{write-debug "Windows PowerShell snap-in: $snapin - Already Loaded"}
    }#End For
    
    if ($FoundMissingSnapin){
        write-warning "Missing Windows PowerShell snap-ins Detected:"
        $missingSnapins | % {write-warning "($_)"}
        return $False
    }#End If
    
    Else{
        Return $true
    }#End Else
    
}#End Function

Function WriteWordLine
#function created by Ryan Revord
#@rsrevord on Twitter
#function created to make output to Word easy in this script
{
	Param( [int]$style=0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [switch]$nonewline)
	$output=""
	#Build output style
	switch ($style)
	{
		0 {$Selection.Style = "No Spacing"}
		1 {$Selection.Style = "Heading 1"}
		2 {$Selection.Style = "Heading 2"}
		3 {$Selection.Style = "Heading 3"}
		Default {$Selection.Style = "No Spacing"}
	}
	#build # of tabs
	While( $tabs -gt 0 ) { 
		$output += "`t"; $tabs--; 
	}
		
	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
	
	#test for new WriteWordLine 0.
	If($nonewline){
		# Do nothing.
	} Else {
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop=$properties | foreach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$null,$_,$null)
		if ($propname -eq $Name) 
		{
			Return $_
		}
	} #foreach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$null,$prop,$Value)
}

Function Process2003Policies
{
	#HDX 3D
	If($Setting.ImageAccelerationState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "HDX 3D\Progressive Display\"
		WriteWordLine 0 3 "Progressive Display: " $Setting.ImageAccelerationState
		If($Setting.ImageAccelerationState -eq "Enabled")
		{
			WriteWordLine 0 3 "Compression level: " -nonewline
			
			switch ($Setting.ImageAccelerationCompressionLevel)
			{
				"HighCompression"   {WriteWordLine 0 0 "High compression; lower image quality"}
				"MediumCompression" {WriteWordLine 0 0 "Medium compression; good image quality"}
				"LowCompression"    {WriteWordLine 0 0 "Low compression; best image quality"}
				"NoCompression"     {WriteWordLine 0 0 "Do not use lossy compression"}
				Default {WriteWordLine 0 0 "Compression level could not be determined: $($Setting.ImageAccelerationCompressionLevel)"}
			}
			If($Setting.ImageAccelerationCompressionIsRestricted)
			{
				WriteWordLine 0 3 "Restrict compression to connections under this "
				WriteWordLine 0 4 "bandwidth\Threshold (Kb/sec): " $Setting.ImageAccelerationCompressionLimit	
			}
			WriteWordLine 0 3 "SpeedScreen Progressive Display compression level: "
			switch ($Setting.ImageAccelerationProgressiveLevel)
			{
				"UltrahighCompression" {WriteWordLine 0 4 "Ultra high compression; ultra low quality"}
				"VeryHighCompression"  {WriteWordLine 0 4 "Very high compression; very low quality"}
				"HighCompression"      {WriteWordLine 0 4 "High compression; low quality"}
				"MediumCompression"    {WriteWordLine 0 4 "Medium compression; medium quality"}
				"LowCompression"       {WriteWordLine 0 4 "Low compression; high quality"}
				"Disabled"             {WriteWordLine 0 4 "Disabled; no progressive display"}
				Default {WriteWordLine 0 0 "SpeedScreen Progressive Display compression level could not be determined: $($Setting.ImageAccelerationProgressiveLevel)"}
			}
			If($Setting.ImageAccelerationProgressiveIsRestricted)
			{
				WriteWordLine 0 3 "Restrict compression to connections under this "
				WriteWordLine 0 4 "bandwidth\Threshold (Kb/sec): " $Setting.ImageAccelerationProgressiveLimit	
			}
			WriteWordLine 0 3 "Use Heavyweight compression (extra CPU, retains quality): " -nonewline
			If($Setting.ImageAccelerationIsHeavyweightUsed)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
	}
	
	#HDX Broadcast
	$xArray = ($Setting.TurnWallpaperOffState, $Setting.TurnWindowContentsOffState )
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "HDX Broadcast\Visual Effects\"
		If($Setting.TurnWallpaperOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn off desktop wallpaper: " $Setting.TurnWallpaperOffState
		}
			
		If($Setting.TurnWindowContentsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn off window contents while dragging: " $Setting.TurnWindowContentsOffState
		}
	}
	$xArray = (	$Setting.SessionAudioState, 		$Setting.SessionClipboardState,
			$Setting.SessionComportsState, 	$Setting.SessionDrivesState,
			$Setting.SessionLptPortsState, 	$Setting.SessionOemChannelsState,
			$Setting.SessionOverallState, 	$Setting.SessionPrinterBandwidthState,
			$Setting.SessionTwainRedirectionState )
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "HDX Broadcast\Session Limits\"
		If($Setting.SessionAudioState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio: " $Setting.SessionAudioState
			If($Setting.SessionAudioState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionAudioLimit
			}
		}
		If($Setting.SessionClipboardState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Clipboard: " $Setting.SessionClipboardState
			If($Setting.SessionClipboardState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionClipboardLimit
			}
		}
		If($Setting.SessionComportsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "COM Ports: " $Setting.SessionComportsState
			If($Setting.SessionComportsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionComportsLimit
			}
		}
		If($Setting.SessionDrivesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives: " $Setting.SessionDrivesState
			If($Setting.SessionDrivesState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionDrivesLimit
			}
		}
		If($Setting.SessionLptPortsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "LPT Ports: " $Setting.SessionLptPortsState
			If($Setting.SessionLptPortsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionLptPortsLimit
			}
		}
		If($Setting.SessionOemChannelsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "OEM Virtual Channels: " $Setting.SessionOemChannelsState
			If($Setting.SessionOemChannelsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionOemChannelsLimit
			}
		}
		If($Setting.SessionOverallState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Overall Session: " $Setting.SessionOverallState
			If($Setting.SessionOverallState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionOverallLimit
			}
		}
		If($Setting.SessionPrinterBandwidthState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer: " $Setting.SessionPrinterBandwidthState
			If($Setting.SessionPrinterBandwidthState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionPrinterBandwidthLimit
			}
		}
		If($Setting.SessionTwainRedirectionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "TWAIN Redirection: " $Setting.SessionTwainRedirectionState
			If($Setting.SessionTwainRedirectionState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionTwainRedirectionLimit
			}
		}
	}

	#HDX Plug-n-Play
	$xArray = (	$Setting.ClientMicrophonesState, 		$Setting.ClientSoundQualityState, 			$Setting.TurnClientAudioMappingOffState,
			$Setting.ClientDrivesState, 			$Setting.ClientDriveMappingState, 			$Setting.ClientAsynchronousWritesState,
			$Setting.TurnComPortsOffState, 		$Setting.TurnLptPortsOffState, 			$Setting.TurnVirtualComPortMappingOffState,
			$Setting.TwainRedirectionState, 		$Setting.TurnClipboardMappingOffState, 		$Setting.TurnOemVirtualChannelsOffState,
			$Setting.TurnAutoClientUpdateOffState, 	$Setting.ClientPrinterAutoCreationState, 		$Setting.LegacyClientPrintersState,
			$Setting.PrinterPropertiesRetentionState,	$Setting.PrinterJobRoutingState, 			$Setting.TurnClientPrinterMappingOffState,
			$Setting.DriverAutoInstallState, 		$Setting.UniversalDriverState, 			$Setting.SessionPrintersState,
			$Setting.ContentRedirectionState, 		$Setting.TurnClientLocalTimeEstimationOffState,	$Setting.TurnClientLocalTimeEstimationOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "HDX Plug-n-Play\Client Resources\"
		If($Setting.ClientMicrophonesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Microphones: " $Setting.ClientMicrophonesState
			If($Setting.ClientMicrophonesState -eq "Enabled")
			{
				If($Setting.ClientMicrophonesAreUsed)
				{
					WriteWordLine 0 4 "Use client microphones for audio input"
				}
				Else
				{
					WriteWordLine 0 4 "Do not use client microphones for audio input"
				}
			}
		}
		If($Setting.ClientSoundQualityState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Sound quality: " $Setting.ClientSoundQualityState
			If($Setting.ClientSoundQualityState)
			{
				WriteWordLine 0 4 "Maximum allowable client audio quality: " 
				switch ($Setting.ClientSoundQualityLevel)
				{
					"Medium" {WriteWordLine 0 5 "Optimized for Speech"}
					"Low"    {WriteWordLine 0 5 "Low Bandwidth"}
					"High"   {WriteWordLine 0 5 "High Definition"}
					Default {WriteWordLine 0 0 "Maximum allowable client audio quality could not be determined: $($Setting.ClientSoundQualityLevel)"}
				}
			}
		}
		If($Setting.TurnClientAudioMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Turn off speakers: " $Setting.TurnClientAudioMappingOffState
			If($Setting.TurnClientAudioMappingOffState -eq "Enabled")
			{
				WriteWordLine 0 4 "Turn off audio mapping to client speakers"
			}
		}
		If($Setting.ClientDrivesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives\Connection: " $Setting.ClientDrivesState
			If($Setting.ClientDrivesState -eq "Enabled")
			{
				If($Setting.ClientDrivesAreConnected)
				{
					WriteWordLine 0 4 "Connect Client Drives at Logon"
				}
				Else
				{
					WriteWordLine 0 4 "Do Not Connect Client Drives at Logon"
				}
			}
		}
		If($Setting.ClientDriveMappingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives\Mappings: " $Setting.ClientDriveMappingState
			If($Setting.ClientDriveMappingState -eq "Enabled")
			{
				If($Setting.TurnFloppyDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Floppy disk drives"	
				}
				If($Setting.TurnHardDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Hard drives"	
				}
				If($Setting.TurnCDRomDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off CD-ROM drives"	
				}
				If($Setting.TurnRemoteDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Remote drives"	
				}
			}
		}
		If($Setting.ClientAsynchronousWritesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives\Optimize\Asynchronous writes: " $Settings.ClientAsynchronousWritesState
			If($Setting.ClientAsynchronousWritesState -eq "Enabled")
			{
				WriteWordLine 0 4 "Turn on asynchronous disk writes to client disks"
			}
		}

		If($Setting.TurnComPortsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Ports\Turn off COM ports: " $Setting.TurnComPortsOffState
		}
		If($Setting.TurnLptPortsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Ports\Turn off LPT ports: " $Setting.TurnLptPortsOffState
		}
		If($Setting.TurnVirtualComPortMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "PDA Devices\Turn on automatic virtual COM port mapping: " $Setting.TurnVirtualComPortMappingOffState
		}
		If($Setting.TwainRedirectionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Other\Configure TWAIN redirection: " $Setting.TwainRedirectionState
			If($Setting.TwainRedirectionState -eq "Enabled")
			{
				If($Setting.TwainRedirectionAllowed)
				{
					WriteWordLine 0 4 "Allow TWAIN redirection"
					If($Setting.TwainRedirectionImageCompression -eq "NoCompression")
					{
						WriteWordLine 0 4 "Do not use lossy compression for high color images"
					}
					Else
					{
						WriteWordLine 0 4 "Use lossy compression for high color images: "
						
						switch ($Setting.TwainRedirectionImageCompression)
						{
							"HighCompression"   {WriteWordLine 0 5 "High compression; lower image quality"}
							"MediumCompression" {WriteWordLine 0 5 "Medium compression; good image quality"}
							"LowCompression"    {WriteWordLine 0 5 "Low compression; best image quality"}
							Default {WriteWordLine 0 0 "Lossy compression for high color images could not be determined: $($Setting.TwainRedirectionImageCompression)"}
						}
					}
				}
				Else
				{
					WriteWordLine 0 4 "Do not allow TWAIN redirection"
				}
			}
		}
		If($Setting.TurnClipboardMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Other\Turn off clipboard mapping: " $Setting.TurnClipboardMappingOffState
		}
		If($Setting.TurnOemVirtualChannelsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Other\Turn off OEM virtual channels: " $Setting.TurnOemVirtualChannelsOffState
		}
	
		If($Setting.TurnAutoClientUpdateOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Client Maintenance\Turn off auto client update: " $Setting.TurnAutoClientUpdateOffState
		}
		If($Setting.ClientPrinterAutoCreationState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Client Printers\Auto-creation: " $Setting.ClientPrinterAutoCreationState
			If($Setting.ClientPrinterAutoCreationState -eq "Enabled")
			{
				WriteWordLine 0 4 "When connecting:"
				switch ($Setting.ClientPrinterAutoCreationOption)
				{
					"LocalPrintersOnly"  {WriteWordLine 0 5 "Auto-create local (non-network) client printers only"}
					"AllPrinters"        {WriteWordLine 0 5 "Auto-create all client printers"}
					"DefaultPrinterOnly" {WriteWordLine 0 5 "Auto-create the client's default printer only"}
					"DoNotAutoCreate"    {WriteWordLine 0 5 "Do not auto-create client printers"}
					Default {WriteWordLine 0 0 "Client Printers\Auto-creation could not be determined: $($Setting.ClientPrinterAutoCreationOption)"}
				}
			}
		}

		If($Setting.LegacyClientPrintersState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Client Printers\Legacy client printers: " $Setting.LegacyClientPrintersState
			If($Setting.LegacyClientPrintersState -eq "Enabled")
			{
				If($Setting.LegacyClientPrintersDynamic)
				{
					WriteWordLine 0 4 "Create dynamic session-private client printers"
				}
				Else
				{
					WriteWordLine 0 4 "Create old-style client printers"
				}
			}
		}
		If($Setting.PrinterPropertiesRetentionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Client Printers\Printer properties retention: " $Setting.PrinterPropertiesRetentionState
			If($Setting.PrinterPropertiesRetentionState -eq "Enabled")
			{
				WriteWordLine 0 4 "Printer properties " -nonewline
				
				switch ($Setting.PrinterPropertiesRetentionOption)
				{
					"FallbackToProfile"     {WriteWordLine 0 0 "Held in profile only if not saved on client"}
					"RetainedInUserProfile" {WriteWordLine 0 0 "Retained in user profile only"}
					"SavedOnClientDevice"   {WriteWordLine 0 0 "Saved on the client device only"}
					Default {WriteWordLine 0 0 "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetentionOption)"}
				}
			}
		}
		If($Setting.PrinterJobRoutingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Client Printers\Print job routing: " $Setting.PrinterJobRoutingState
			If($Setting.PrinterJobRoutingState -eq "Enabled")
			{
				WriteWordLine 0 4 "For client printers on a network printer server: "
				If($Setting.PrinterJobRoutingDirect)
				{
					WriteWordLine 0 5 "Connect directly to network print server if possible"
				}
				Else
				{
					WriteWordLine 0 5 "Always connect indirectly as a client printer"
				}
			}
		}
		If($Setting.TurnClientPrinterMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Client Printers\Turn off client printer mapping: " $Setting.TurnClientPrinterMappingOffState
		}
		If($Setting.DriverAutoInstallState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Drivers\Native printer driver auto-install: " $Setting.DriverAutoInstallState
			If($Setting.DriverAutoInstallState -eq "Enabled")
			{
				If($Setting.DriverAutoInstallAsNeeded)
				{
					WriteWordLine 0 4 "Install Windows native drivers as needed"
				}
				Else
				{
					WriteWordLine 0 4 "Do not automatically install drivers"
				}
			}
		}
		If($Setting.UniversalDriverState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Drivers\Universal driver: " $Setting.UniversalDriverState
			If($Setting.UniversalDriverState -eq "Enabled")
			{
				WriteWordLine 0 4 "When auto-creating client printers: "
				
				switch ($Setting.UniversalDriverOption)
				{
					"FallbackOnly"  {WriteWordLine 0 4 "Use universal driver only if requested driver is unavailable"}
					"SpecificOnly"  {WriteWordLine 0 4 "Use only printer model specific drivers"}
					"ExclusiveOnly" {WriteWordLine 0 4 "Use universal driver only"}
					Default {WriteWordLine 0 0 "When auto-creating client printers could not be determined: $($Setting.UniversalDriverOption)"}
				}
			}
		}
		If($Setting.SessionPrintersState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printing\Session printers\Session printers: " $Setting.SessionPrintersState
			If($Setting.SessionPrintersState -eq "Enabled")
			{
				If($Setting.SessionPrinterList)
				{
					WriteWordLine 0 4 "Network printers to connect at logon:"
					ForEach($Printer in $Setting.SessionPrinterList)
					{
						WriteWordLine 0 5 $Printer
						$index = $Printer.SubString( 2 ).IndexOf( '\' )
						if( $index -ge 0 )
						{
							$srv = $Printer.SubString( 0, $index + 2 )
							$ptr  = $Printer.SubString( $index + 3 )
						}
						$SessionPrinterSettings = Get-XASessionPrinter -PolicyName $Policy.PolicyName -PrinterName $Ptr -EA 0
						If(![String]::IsNullOrEmpty( $SessionPrinterSettings ))
						{
							If($SessionPrinterSettings.ApplyCustomSettings)
							{
								WriteWordLine 0 5 "Shared Name`t: " $SessionPrinterSettings.PrinterName
								WriteWordLine 0 5 "Server`t`t: " $SessionPrinterSettings.ServerName
								WriteWordLine 0 5 "Printer Model`t: " $SessionPrinterSettings.DriverName
								WriteWordLine 0 5 "Location`t: " $SessionPrinterSettings.Location
								WriteWordLine 0 5 "Paper Size`t: " -nonewline
								switch ($SessionPrinterSettings.PaperSize)
								{
									"A4"          {WriteWordLine 0 0 "A4"}
									"A4Small"     {WriteWordLine 0 0 "A4 Small"}
									"Envelope10"  {WriteWordLine 0 0 "Envelope #10"}
									"EnvelopeB5"  {WriteWordLine 0 0 "Envelope B5"}
									"EnvelopeC5"  {WriteWordLine 0 0 "Envelope C5"}
									"EnvelopeDL"  {WriteWordLine 0 0 "Envelope DL"}
									"Monarch"     {WriteWordLine 0 0 "Envelope Monarch"}
									"Executive"   {WriteWordLine 0 0 "Executive"}
									"Legal"       {WriteWordLine 0 0 "Legal"}
									"Letter"      {WriteWordLine 0 0 "Letter"}
									"LetterSmall" {WriteWordLine 0 0 "Letter Small"}
									"Note" {WriteWordLine 0 0 "Note"}
									Default 
									{
										WriteWordLine 0 0 "Custom..."
										WriteWordLine 0 5 "Width`t`t: $($SessionPrinterSettings.Width) (Millimeters)" 
										WriteWordLine 0 5 "Height`t`t: $($SessionPrinterSettings.Height) (Millimeters)" 
									}
								}
								WriteWordLine 0 5 "Copy Count`t: " $SessionPrinterSettings.CopyCount
								If($SessionPrinterSettings.CopyCount -gt 1)
								{
									WriteWordLine 0 5 "Collated`t: " -nonewline
									If($SessionPrinterSettings.Collated)
									{
										WriteWordLine 0 0 "Yes"
									}
									Else
									{
										WriteWordLine 0 0 "No"
									}
								}
								WriteWordLine 0 5 "Print Quality`t: " -nonewline
								switch ($SessionPrinterSettings.PrintQuality)
								{
									"Dpi600" {WriteWordLine 0 0 "600 dpi"}
									"Dpi300" {WriteWordLine 0 0 "300 dpi"}
									"Dpi150" {WriteWordLine 0 0 "150 dpi"}
									"Dpi75"  {WriteWordLine 0 0 "75 dpi"}
									Default {WriteWordLine 0 0 "Print Quality could not be determined: $($SessionPrinterSettings.PrintQuality)"}
								}
								WriteWordLine 0 5 "Orientation`t: " $SessionPrinterSettings.PaperOrientation
								WriteWordLine 0 5 "Apply customized settings at every logon: " -nonewline
								If($SessionPrinterSettings.ApplySettingsOnLogOn)
								{
									WriteWordLine 0 0 "Yes"
								}
								Else
								{
									WriteWordLine 0 0 "No"
								}
							}
						}
					}
				}
				WriteWordLine 0 3 "Printing\Session printers\Client's default printer: "
				If($Setting.SessionPrinterDefaultOption -eq "SetToPrinterIndex")
				{
					WriteWordLine 0 4 $Setting.SessionPrinterList[$Setting.SessionPrinterDefaultIndex]
				}
				Else
				{
					switch ($Setting.SessionPrinterDefaultOption)
					{
						"SetToClientMainPrinter" {WriteWordLine 0 4 "Set default printer to the client's main printer"}
						"DoNotAdjust"            {WriteWordLine 0 4 "Do not adjust the user's default printer"}
						Default {WriteWordLine 0 0 "Client's default printer could not be determined: $($Setting.SessionPrinterDefaultOption)"}
					}
					
				}
			}
		}
		If($Setting.ContentRedirectionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Content Redirection\Server to client: " $Setting.ContentRedirectionState
			If($Setting.ContentRedirectionState -eq "Enabled")
			{
				If($Setting.ContentRedirectionIsUsed)
				{
					WriteWordLine 0 4 "Use Content Redirection from server to client"
				}
				Else
				{
					WriteWordLine 0 4 "Do not use Content Redirection from server to client"
				}
			}
		}
		If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Time Zones\Do Not Estimate Client Local Time: " $Setting.TurnClientLocalTimeEstimationOffState
		}
		If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Time Zones\Do Not Use Client's Local Time: " $Setting.TurnClientLocalTimeOffState
		}
	}
	
	#User Workspace
	$xArray = (	$Setting.ConcurrentSessionsState, 	$Setting.ZonePreferenceAndFailoverState, 	$Setting.ShadowingState, `
			$Setting.ShadowingPermissionsState,	$Setting.CentralCredentialStoreState, 	$Setting.TurnPasswordManagerOffState, `
			$Setting.StreamingDeliveryProtocolState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "User Workspace\"
		If($Setting.ConcurrentSessionsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Connections\Limit total concurrent sessions: " $Setting.ConcurrentSessionsState
			If($Setting.ConcurrentSessionsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit: " $Setting.ConcurrentSessionsLimit
			}
		}
		If($Setting.ZonePreferenceAndFailoverState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Connections\Zone preference and failover: " $Setting.ZonePreferenceAndFailoverState
			If($Setting.ZonePreferenceAndFailoverState -eq "Enabled")
			{
				WriteWordLine 0 4 "Zone preference settings:"
				ForEach($Pref in $Setting.ZonePreferences)
				{
					WriteWordLine 0 5 $Pref
				}
			}
		}
		If($Setting.ShadowingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Shadowing\Configuration: " $Setting.ShadowingState
			If($Setting.ShadowingState -eq "Enabled")
			{
				If($Setting.ShadowingAllowed)
				{
					WriteWordLine 0 4 "Allow Shadowing"
					WriteWordLine 0 4 "Prohibit Being Shadowed Without Notification: " $Setting.ShadowingProhibitedWithoutNotification
					WriteWordLine 0 4 "Prohibit Remote Input When Being Shadowed: " $Setting.ShadowingRemoteInputProhibited
				}
				Else
				{
					WriteWordLine 0 2 "Do Not Allow Shadowing"
				}
			}
		}
		If($Setting.ShadowingPermissionsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Shadowing\Permissions: " $Setting.ShadowingPermissionsState
			If($Setting.ShadowingPermissionsState -eq "Enabled")
			{
				If($Setting.ShadowingAccountsAllowed)
				{
					WriteWordLine 0 4 "Accounts allowed to shadow:"
					ForEach($Allowed in $Setting.ShadowingAccountsAllowed)
					{
						WriteWordLine 0 5 $Allowed
					}
				}
				If($Setting.ShadowingAccountsDenied)
				{
					WriteWordLine 0 4 "Accounts denied from shadowing:"
					ForEach($Denied in $Setting.ShadowingAccountsDenied)
					{
						WriteWordLine 0 5 $Denied
					}
				}
			}
		}
		If($Setting.CentralCredentialStoreState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Single Sign-On\Central Credential Store: " $Setting.CentralCredentialStoreState
			If($Setting.CentralCredentialStoreState -eq "Enabled")
			{
				If($Setting.CentralCredentialStorePath)
				{
					WriteWordLine 0 4 "UNC path of Central Credential Store: " $Setting.CentralCredentialStorePath
				}
				Else
				{
					WriteWordLine 0 4 "No UNC path to Central Credential Store entered"
				}
			}
		}
		If($Setting.TurnPasswordManagerOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Single Sign-On\Do not use Citrix Password Manager: " $Setting.TurnPasswordManagerOffState
		}
		If($Setting.StreamingDeliveryProtocolState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Streamed Applications\Configure delivery protocol: " $Setting.StreamingDeliveryProtocolState
			If($Setting.StreamingDeliveryProtocolState -eq "Enabled")
			{
				WriteWordLine 0 4 "Streaming Delivery Protocol option: " 
				switch ($Setting.StreamingDeliveryOption)
				{
					"Unknown"                {WriteWordLine 0 5 "Unknown"}
					"ForceServerAccess"      {WriteWordLine 0 5 "Do not allow applications to stream to the client"}
					"ForcedStreamedDelivery" {WriteWordLine 0 5 "Force applications to stream to the client"}
					Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
				}
			}
		}
	}

	#Security
	If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "Security\Encryption\SecureICA encryption: " $Setting.SecureIcaEncriptionState
		If($Setting.SecureIcaEncriptionState -eq "Enabled")
		{
			WriteWordLine 0 3 "Encryption level: " -nonewline
			switch ($Setting.SecureIcaEncriptionLevel)
			{
				"Unknown" {WriteWordLine 0 0 "Unknown encryption"}
				"Basic"   {WriteWordLine 0 0 "Basic"}
				"LogOn"   {WriteWordLine 0 0 "RC5 (128 bit) logon only"}
				"Bits40"  {WriteWordLine 0 0 "RC5 (40 bit)"}
				"Bits56"  {WriteWordLine 0 0 "RC5 (56 bit)"}
				"Bits128" {WriteWordLine 0 0 "RC5 (128 bit)"}
				Default {WriteWordLine 0 0 "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"}
			}
		}
	}
}

Function Process2008Policies
{
	#Bandwidth
	$xArray = ($Setting.TurnWallpaperOffState, $Setting.TurnWindowContentsOffState, $Setting.TurnWindowContentsOffState )
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "Bandwidth\Visual Effects\"
		If($Setting.TurnWallpaperOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn Off Desktop Wallpaper: " $Setting.TurnWallpaperOffState
		}
		If($Setting.TurnMenuAnimationsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn Off Menu and Windows Animations: " $Setting.TurnMenuAnimationsOffState
		}
		If($Setting.TurnWindowContentsOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn Off Window Contents While Dragging: " $Setting.TurnWindowContentsOffState
		}
	}
	
	If($Setting.ImageAccelerationState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "Bandwidth\SpeedScreen\"
		WriteWordLine 0 3 "Image acceleration using lossy compression: " $Setting.ImageAccelerationState
		If($Setting.ImageAccelerationState -eq "Enabled")
		{
			WriteWordLine 0 3 "Compression level: " -nonewline
			
			switch ($Setting.ImageAccelerationCompressionLevel)
			{
				"HighCompression"   {WriteWordLine 0 0 "High compression; lower image quality"}
				"MediumCompression" {WriteWordLine 0 0 "Medium compression; good image quality"}
				"LowCompression"    {WriteWordLine 0 0 "Low compression; best image quality"}
				"NoCompression"     {WriteWordLine 0 0 "Do not use lossy compression"}
				Default {WriteWordLine 0 0 "Compression level could not be determined: $($Setting.ImageAccelerationCompressionLevel)"}
			}
			If($Setting.ImageAccelerationCompressionIsRestricted)
			{
				WriteWordLine 0 3 "Restrict compression to connections under this "
				WriteWordLine 0 4 "bandwidth\Threshold (Kb/sec): " $Setting.ImageAccelerationCompressionLimit	
			}
			WriteWordLine 0 3 "SpeedScreen Progressive Display compression level: "
			switch ($Setting.ImageAccelerationProgressiveLevel)
			{
				"UltrahighCompression" {WriteWordLine 0 4 "Ultra high compression; ultra low quality"}
				"VeryHighCompression"  {WriteWordLine 0 4 "Very high compression; very low quality"}
				"HighCompression"      {WriteWordLine 0 4 "High compression; low quality"}
				"MediumCompression"    {WriteWordLine 0 4 "Medium compression; medium quality"}
				"LowCompression"       {WriteWordLine 0 4 "Low compression; high quality"}
				"Disabled"             {WriteWordLine 0 4 "Disabled; no progressive display"}
				Default {WriteWordLine 0 0 "SpeedScreen Progressive Display compression level could not be determined: $($Setting.ImageAccelerationProgressiveLevel)"}
			}
			If($Setting.ImageAccelerationProgressiveIsRestricted)
			{
				WriteWordLine 0 3 "Restrict compression to connections under this "
				WriteWordLine 0 4 "bandwidth\Threshold (Kb/sec): " $Setting.ImageAccelerationProgressiveLimit	
			}
			WriteWordLine 0 3 "Use Heavyweight compression (extra CPU, retains quality): " -nonewline
			If($Setting.ImageAccelerationIsHeavyweightUsed)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
	}
	
	$xArray = (	$Setting.SessionAudioState,	$Setting.SessionClipboardState,		$Setting.SessionComportsState, 
			$Setting.SessionDrivesState,	$Setting.SessionLptPortsState,		$Setting.SessionOemChannelsState, 
			$Setting.SessionOverallState,	$Setting.SessionPrinterBandwidthState,	$Setting.SessionTwainRedirectionState )
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "Bandwidth\Session Limits\"
		If($Setting.SessionAudioState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio: " $Setting.SessionAudioState
			If($Setting.SessionAudioState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionAudioLimit
			}
		}
		If($Setting.SessionClipboardState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Clipboard: " $Setting.SessionClipboardState
			If($Setting.SessionClipboardState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionClipboardLimit
			}
		}
		If($Setting.SessionComportsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "COM Ports: " $Setting.SessionComportsState
			If($Setting.SessionComportsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionComportsLimit
			}
		}
		If($Setting.SessionDrivesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives: " $Setting.SessionDrivesState
			If($Setting.SessionDrivesState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionDrivesLimit
			}
		}
		If($Setting.SessionLptPortsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "LPT Ports: " $Setting.SessionLptPortsState
			If($Setting.SessionLptPortsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionLptPortsLimit
			}
		}
		If($Setting.SessionOemChannelsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "OEM Virtual Channels: " $Setting.SessionOemChannelsState
			If($Setting.SessionOemChannelsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionOemChannelsLimit
			}
		}
		If($Setting.SessionOverallState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Overall Session: " $Setting.SessionOverallState
			If($Setting.SessionOverallState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionOverallLimit
			}
		}
		If($Setting.SessionPrinterBandwidthState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer: " $Setting.SessionPrinterBandwidthState
			If($Setting.SessionPrinterBandwidthState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionPrinterBandwidthLimit
			}
		}
		If($Setting.SessionTwainRedirectionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "TWAIN Redirection: " $Setting.SessionTwainRedirectionState
			If($Setting.SessionTwainRedirectionState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit (Kb/sec): " $Setting.SessionTwainRedirectionLimit
			}
		}
	}

	$xArray = (	$Setting.SessionAudioPercentState,		$Setting.SessionClipboardPercentState,		$Setting.SessionComportsPercentState, 
			$Setting.SessionDrivesPercentState,		$Setting.SessionLptPortsPercentState,		$Setting.SessionOemChannelsPercentState, 
			$Setting.SessionOverallState,	$Setting.SessionPrinterPercentState,	$Setting.SessionTwainRedirectionPercentState )
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 'Bandwidth\Session Limits (%)\'
		If($Setting.SessionAudioPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 'Audio: ' $Setting.SessionAudioPercentState
			If($Setting.SessionAudioPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionAudioPercentLimit
			}
		}
		If($Setting.SessionClipboardPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Clipboard: " $Setting.SessionClipboardPercentState
			If($Setting.SessionClipboardPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionClipboardPercentLimit
			}
		}
		If($Setting.SessionComportsPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "COM Ports: " $Setting.SessionComportsPercentState
			If($Setting.SessionComportsPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionComportsPercentLimit
			}
		}
		If($Setting.SessionDrivesPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Drives: " $Setting.SessionDrivesPercentState
			If($Setting.SessionDrivesPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionDrivesPercentLimit
			}
		}
		If($Setting.SessionLptPortsPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "LPT Ports: " $Setting.SessionLptPortsPercentState
			If($Setting.SessionLptPortsPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionLptPortsPercentLimit
			}
		}
		If($Setting.SessionOemChannelsPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "OEM Virtual Channels: " $Setting.SessionOemChannelsPercentState
			If($Setting.SessionOemChannelsPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionOemChannelsPercentLimit
			}
		}
		If($Setting.SessionPrinterPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer: " $Setting.SessionPrinterPercentState
			If($Setting.SessionPrinterPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionPrinterPercentLimit
			}
		}
		If($Setting.SessionTwainRedirectionPercentState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "TWAIN Redirection: " $Setting.SessionTwainRedirectionPercentState
			If($Setting.SessionTwainRedirectionPercentState -eq "Enabled")
			{
				WriteWordLine 0 4 'Limit (%): ' $Setting.SessionTwainRedirectionPercentLimit
			}
		}
	}
	
	$xArray = (	$Setting.ClientMicrophonesState,	$Setting.ClientSoundQualityState,		$Setting.TurnClientAudioMappingOffState,
			$Setting.ClientDrivesState,		$Setting.ClientDriveMappingState,		$Setting.ClientAsynchronousWritesState,
			$Setting.TwainRedirectionState,	$Setting.TurnClipboardMappingOffState,	$Setting.TurnOemVirtualChannelsOffState,
			$Setting.TurnComPortsOffState,	$Setting.TurnLptPortsOffState,		$Setting.TurnVirtualComPortMappingOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "Client Devices\Resources"
		If($Setting.ClientMicrophonesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Microphones: " $Setting.ClientMicrophonesState
			If($Setting.ClientMicrophonesState -eq "Enabled")
			{
				If($Setting.ClientMicrophonesAreUsed)
				{
					WriteWordLine 0 4 "Use client microphones for audio input"
				}
				Else
				{
					WriteWordLine 0 4 "Do not use client microphones for audio input"
				}
			}
		}
		If($Setting.ClientSoundQualityState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Sound quality: " $Setting.ClientSoundQualityState
			If($Setting.ClientSoundQualityState)
			{
				WriteWordLine 0 4 "Maximum allowable client audio quality: " 
				switch ($Setting.ClientSoundQualityLevel)
				{
					"Medium" {WriteWordLine 0 5 "Optimized for Speech"}
					"Low"    {WriteWordLine 0 5 "Low Bandwidth"}
					"High"   {WriteWordLine 0 5 "High Definition"}
					Default {WriteWordLine 0 0 "Maximum allowable client audio quality could not be determined: $($Setting.ClientSoundQualityLevel)"}
				}
			}
		}
		If($Setting.TurnClientAudioMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Audio\Turn off speakers: " $Setting.TurnClientAudioMappingOffState
			If($Setting.TurnClientAudioMappingOffState -eq "Enabled")
			{
				WriteWordLine 0 4 "Turn off audio mapping to client speakers"
			}
		}

		WriteWordLine 0 2 "Client Devices\Resources\Drives"
		If($Setting.ClientDrivesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Connection: " $Setting.ClientDrivesState
			If($Setting.ClientDrivesState -eq "Enabled")
			{
				If($Setting.ClientDrivesAreConnected)
				{
					WriteWordLine 0 4 "Connect Client Drives at Logon"
				}
				Else
				{
					WriteWordLine 0 4 "Do Not Connect Client Drives at Logon"
				}
			}
		}
		If($Setting.ClientDriveMappingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Mappings: " $Setting.ClientDriveMappingState
			If($Setting.ClientDriveMappingState -eq "Enabled")
			{
				If($Setting.TurnFloppyDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Floppy disk drives"	
				}
				If($Setting.TurnHardDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Hard drives"	
				}
				If($Setting.TurnCDRomDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off CD-ROM drives"	
				}
				If($Setting.TurnRemoteDriveMappingOff)
				{
					WriteWordLine 0 4 "Turn off Remote drives"	
				}
			}
		}
		If($Setting.ClientAsynchronousWritesState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Optimize\Asynchronous writes: " $Setting.ClientAsynchronousWritesState
			If($Setting.ClientAsynchronousWritesState -eq "Enabled")
			{
				WriteWordLine 0 4 "Turn on asynchronous disk writes to client disks"
			}

			WriteWordLine 0 3 "Special folder redirection: " $Setting.TurnSpecialFolderRedirectionOffState
			If($Setting.TurnSpecialFolderRedirectionOffState -eq "Enabled")
			{
				WriteWordLine 0 4 "Do not allow special folder redirection"
			}
		}

		$xArray = ($Setting.TwainRedirectionState, $Setting.TurnClipboardMappingOffState, $Setting.TurnOemVirtualChannelsOffState)
		If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
		{
			WriteWordLine 0 2 "Client Devices\Resources\Other"
			If($Setting.TwainRedirectionState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Configure TWAIN redirection: " $Setting.TwainRedirectionState
				If($Setting.TwainRedirectionState -eq "Enabled")
				{
					If($Setting.TwainRedirectionAllowed)
					{
						WriteWordLine 0 4 "Allow TWAIN redirection"
						If($Setting.TwainRedirectionImageCompression -eq "NoCompression")
						{
							WriteWordLine 0 4 "Do not use lossy compression for high color images"
						}
						Else
						{
							WriteWordLine 0 4 "Use lossy compression for high color images: "
							
							switch ($Setting.TwainRedirectionImageCompression)
							{
								"HighCompression"   {WriteWordLine 0 5 "High compression; lower image quality"}
								"MediumCompression" {WriteWordLine 0 5 "Medium compression; good image quality"}
								"LowCompression"    {WriteWordLine 0 5 "Low compression; best image quality"}
								Default {WriteWordLine 0 0 "Lossy compression for high color images could not be determined: $($Setting.TwainRedirectionImageCompression)"}
							}
						}
					}
					Else
					{
						WriteWordLine 0 4 "Do not allow TWAIN redirection"
					}
				}
			}
			If($Setting.TurnClipboardMappingOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off clipboard mapping: " $Setting.TurnClipboardMappingOffState
			}
			If($Setting.TurnOemVirtualChannelsOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off OEM virtual channels: " $Setting.TurnOemVirtualChannelsOffState
			}
		}

		$xArray = ($Setting.TurnComPortsOffState, $Setting.TurnLptPortsOffState)
		If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
		{
			WriteWordLine 0 2 "Client Devices\Resources\Ports"
			If($Setting.TurnComPortsOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off COM ports: " $Setting.TurnComPortsOffState
			}
			If($Setting.TurnLptPortsOffState -ne "NotConfigured")
			{
				WriteWordLine 0 3 "Turn off LPT ports: " $Setting.TurnLptPortsOffState
			}
		}
		
		If($Setting.TurnVirtualComPortMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 2 "Client Devices\Resources\PDA Devices"
			WriteWordLine 0 3 "Turn on automatic virtual COM port mapping: " $Setting.TurnVirtualComPortMappingOffState
		}
	}
	
	If($Setting.TurnAutoClientUpdateOffState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "Client Devices\Maintenance"
		WriteWordLine 0 3 "Turn off auto client update: " $Setting.TurnAutoClientUpdateOffState
	}
	
	$xArray = (	$Setting.ClientPrinterAutoCreationState,	$Setting.LegacyClientPrintersState,
			$Setting.PrinterPropertiesRetentionState,	$Setting.PrinterJobRoutingState,
			$Setting.TurnClientPrinterMappingOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "Printing\Client Printers"
		If($Setting.ClientPrinterAutoCreationState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Auto-creation: " $Setting.ClientPrinterAutoCreationState
			If($Setting.ClientPrinterAutoCreationState -eq "Enabled")
			{
				WriteWordLine 0 4 "When connecting:"
				switch ($Setting.ClientPrinterAutoCreationOption)
				{
					"LocalPrintersOnly"  {WriteWordLine 0 5 "Auto-create local (non-network) client printers only"}
					"AllPrinters"        {WriteWordLine 0 5 "Auto-create all client printers"}
					"DefaultPrinterOnly" {WriteWordLine 0 5 "Auto-create the client's default printer only"}
					"DoNotAutoCreate"    {WriteWordLine 0 5 "Do not auto-create client printers"}
					Default {WriteWordLine 0 0 "Client Printers\Auto-creation could not be determined: $($Setting.ClientPrinterAutoCreationOption)"}
				}
			}
		}

		If($Setting.LegacyClientPrintersState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Legacy client printers: " $Setting.LegacyClientPrintersState
			If($Setting.LegacyClientPrintersState -eq "Enabled")
			{
				If($Setting.LegacyClientPrintersDynamic)
				{
					WriteWordLine 0 4 "Create dynamic session-private client printers"
				}
				Else
				{
					WriteWordLine 0 4 "Create old-style client printers"
				}
			}
		}
		If($Setting.PrinterPropertiesRetentionState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Printer properties retention: " $Setting.PrinterPropertiesRetentionState
			If($Setting.PrinterPropertiesRetentionState -eq "Enabled")
			{
				WriteWordLine 0 4 "Printer properties " -nonewline
				
				switch ($Setting.PrinterPropertiesRetentionOption)
				{
					"FallbackToProfile"     {WriteWordLine 0 0 "Held in profile only if not saved on client"}
					"RetainedInUserProfile" {WriteWordLine 0 0 "Retained in user profile only"}
					"SavedOnClientDevice"   {WriteWordLine 0 0 "Saved on the client device only"}
					Default {WriteWordLine 0 0 "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetentionOption)"}
				}
			}
		}
		If($Setting.PrinterJobRoutingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Print job routing: " $Setting.PrinterJobRoutingState
			If($Setting.PrinterJobRoutingState -eq "Enabled")
			{
				WriteWordLine 0 4 "For client printers on a network printer server: "
				If($Setting.PrinterJobRoutingDirect)
				{
					WriteWordLine 0 5 "Connect directly to network print server if possible"
				}
				Else
				{
					WriteWordLine 0 5 "Always connect indirectly as a client printer"
				}
			}
		}
		If($Setting.TurnClientPrinterMappingOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Turn off client printer mapping: " $Setting.TurnClientPrinterMappingOffState
		}
	}
	
	$xArray = ($Setting.DriverAutoInstallState, $Setting.UniversalDriverState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "Printing\Drivers"
		If($Setting.DriverAutoInstallState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Native printer driver auto-install: " $Setting.DriverAutoInstallState
			If($Setting.DriverAutoInstallState -eq "Enabled")
			{
				If($Setting.DriverAutoInstallAsNeeded)
				{
					WriteWordLine 0 4 "Install Windows native drivers as needed"
				}
				Else
				{
					WriteWordLine 0 4 "Do not automatically install drivers"
				}
			}
		}
		If($Setting.UniversalDriverState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Universal driver: " $Setting.UniversalDriverState
			If($Setting.UniversalDriverState -eq "Enabled")
			{
				WriteWordLine 0 4 "When auto-creating client printers: "
				
				switch ($Setting.UniversalDriverOption)
				{
					"FallbackOnly"  {WriteWordLine 0 4 "Use universal driver only if requested driver is unavailable"}
					"SpecificOnly"  {WriteWordLine 0 4 "Use only printer model specific drivers"}
					"ExclusiveOnly" {WriteWordLine 0 4 "Use universal driver only"}
					Default {WriteWordLine 0 0 "When auto-creating client printers could not be determined: $($Setting.UniversalDriverOption)"}
				}
			}
		}
	}
	
	If($Setting.SessionPrintersState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "Printing\Session printers"
		WriteWordLine 0 3 "Session printers: " $Setting.SessionPrintersState
		If($Setting.SessionPrintersState -eq "Enabled")
		{
			If($Setting.SessionPrinterList)
			{
				WriteWordLine 0 4 "Network printers to connect at logon:"
				ForEach($Printer in $Setting.SessionPrinterList)
				{
					WriteWordLine 0 5 $Printer
					$index = $Printer.SubString( 2 ).IndexOf( '\' )
					if( $index -ge 0 )
					{
						$srv = $Printer.SubString( 0, $index + 2 )
						$ptr  = $Printer.SubString( $index + 3 )
					}
					$SessionPrinterSettings = Get-XASessionPrinter -PolicyName $Policy.PolicyName -PrinterName $Ptr -EA 0
					If(![String]::IsNullOrEmpty( $SessionPrinterSettings ))
					{
						If($SessionPrinterSettings.ApplyCustomSettings)
						{
							WriteWordLine 0 5 "Shared Name`t: " $SessionPrinterSettings.PrinterName
							WriteWordLine 0 5 "Server`t`t: " $SessionPrinterSettings.ServerName
							WriteWordLine 0 5 "Printer Model`t: " $SessionPrinterSettings.DriverName
							WriteWordLine 0 5 "Location`t: " $SessionPrinterSettings.Location
							WriteWordLine 0 5 "Paper Size`t: " -nonewline
							switch ($SessionPrinterSettings.PaperSize)
							{
								"A4"          {WriteWordLine 0 0 "A4"}
								"A4Small"     {WriteWordLine 0 0 "A4 Small"}
								"Envelope10"  {WriteWordLine 0 0 "Envelope #10"}
								"EnvelopeB5"  {WriteWordLine 0 0 "Envelope B5"}
								"EnvelopeC5"  {WriteWordLine 0 0 "Envelope C5"}
								"EnvelopeDL"  {WriteWordLine 0 0 "Envelope DL"}
								"Monarch"     {WriteWordLine 0 0 "Envelope Monarch"}
								"Executive"   {WriteWordLine 0 0 "Executive"}
								"Legal"       {WriteWordLine 0 0 "Legal"}
								"Letter"      {WriteWordLine 0 0 "Letter"}
								"LetterSmall" {WriteWordLine 0 0 "Letter Small"}
								"Note" {WriteWordLine 0 0 "Note"}
								Default 
								{
									WriteWordLine 0 0 "Custom..."
									WriteWordLine 0 5 "Width`t`t: $($SessionPrinterSettings.Width) (Millimeters)" 
									WriteWordLine 0 5 "Height`t`t: $($SessionPrinterSettings.Height) (Millimeters)" 
								}
							}
							WriteWordLine 0 5 "Copy Count`t: " $SessionPrinterSettings.CopyCount
							If($SessionPrinterSettings.CopyCount -gt 1)
							{
								WriteWordLine 0 5 "Collated`t: " -nonewline
								If($SessionPrinterSettings.Collated)
								{
									WriteWordLine 0 0 "Yes"
								}
								Else
								{
									WriteWordLine 0 0 "No"
								}
							}
							WriteWordLine 0 5 "Print Quality`t: " -nonewline
							switch ($SessionPrinterSettings.PrintQuality)
							{
								"Dpi600" {WriteWordLine 0 0 "600 dpi"}
								"Dpi300" {WriteWordLine 0 0 "300 dpi"}
								"Dpi150" {WriteWordLine 0 0 "150 dpi"}
								"Dpi75"  {WriteWordLine 0 0 "75 dpi"}
								Default {WriteWordLine 0 0 "Print Quality could not be determined: $($SessionPrinterSettings.PrintQuality)"}
							}
							WriteWordLine 0 5 "Orientation`t: " $SessionPrinterSettings.PaperOrientation
							WriteWordLine 0 5 "Apply customized settings at every logon: " -nonewline
							If($SessionPrinterSettings.ApplySettingsOnLogOn)
							{
								WriteWordLine 0 0 "Yes"
							}
							Else
							{
								WriteWordLine 0 0 "No"
							}
						}
					}
				}
			}
			WriteWordLine 0 3 "Client's default printer: "
			If($Setting.SessionPrinterDefaultOption -eq "SetToPrinterIndex")
			{
				WriteWordLine 0 4 $Setting.SessionPrinterList[$Setting.SessionPrinterDefaultIndex]
			}
			Else
			{
				switch ($Setting.SessionPrinterDefaultOption)
				{
					"SetToClientMainPrinter" {WriteWordLine 0 4 "Set default printer to the client's main printer"}
					"DoNotAdjust"            {WriteWordLine 0 4 "Do not adjust the user's default printer"}
					Default {WriteWordLine 0 0 "Client's default printer could not be determined: $($Setting.SessionPrinterDefaultOption)"}
				}
				
			}
		}
	}

	#User Workspace
	$xArray = ($Setting.ConcurrentSessionsState, $Setting.ZonePreferenceAndFailoverState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "User Workspace\Connections"
		If($Setting.ConcurrentSessionsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Limit total concurrent sessions: " $Setting.ConcurrentSessionsState
			If($Setting.ConcurrentSessionsState -eq "Enabled")
			{
				WriteWordLine 0 4 "Limit: " $Setting.ConcurrentSessionsLimit
			}
		}
		If($Setting.ZonePreferenceAndFailoverState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Zone preference and failover: " $Setting.ZonePreferenceAndFailoverState
			If($Setting.ZonePreferenceAndFailoverState -eq "Enabled")
			{
				WriteWordLine 0 4 "Zone preference settings:"
				ForEach($Pref in $Setting.ZonePreferences)
				{
					WriteWordLine 0 5 $Pref
				}
			}
		}
	}
	
	If($Setting.ContentRedirectionState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "User Workspace\Content Redirection"
		WriteWordLine 0 3 "Server to client: " $Setting.ContentRedirectionState
		If($Setting.ContentRedirectionState -eq "Enabled")
		{
			If($Setting.ContentRedirectionIsUsed)
			{
				WriteWordLine 0 4 "Use Content Redirection from server to client"
			}
			Else
			{
				WriteWordLine 0 4 "Do not use Content Redirection from server to client"
			}
		}
	}

	$xArray = ($Setting.ShadowingState, $Setting.ShadowingPermissionsState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "User Workspace\Shadowing"
		If($Setting.ShadowingState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Configuration: " $Setting.ShadowingState
			If($Setting.ShadowingState -eq "Enabled")
			{
				If($Setting.ShadowingAllowed)
				{
					WriteWordLine 0 4 "Allow Shadowing"
					WriteWordLine 0 4 "Prohibit Being Shadowed Without Notification: " $Setting.ShadowingProhibitedWithoutNotification
					WriteWordLine 0 4 "Prohibit Remote Input When Being Shadowed: " $Setting.ShadowingRemoteInputProhibited
				}
				Else
				{
					WriteWordLine 0 3 "Do Not Allow Shadowing"
				}
			}
		}
		If($Setting.ShadowingPermissionsState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Permissions: " $Setting.ShadowingPermissionsState
			If($Setting.ShadowingPermissionsState -eq "Enabled")
			{
				If($Setting.ShadowingAccountsAllowed)
				{
					WriteWordLine 0 4 "Accounts allowed to shadow:"
					ForEach($Allowed in $Setting.ShadowingAccountsAllowed)
					{
						WriteWordLine 0 5 $Allowed
					}
				}
				If($Setting.ShadowingAccountsDenied)
				{
					WriteWordLine 0 4 "Accounts denied from shadowing:"
					ForEach($Denied in $Setting.ShadowingAccountsDenied)
					{
						WriteWordLine 0 5 $Denied
					}
				}
			}
		}
	}

	$xArray = ($Setting.TurnClientLocalTimeEstimationOffState, $Setting.TurnClientLocalTimeEstimationOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "User Workspace\Time Zones"
		If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Do not estimate local time for legacy clients: " $Setting.TurnClientLocalTimeEstimationOffState
		}
		If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Do not use Client's local time: " $Setting.TurnClientLocalTimeOffState
		}
	}
	
	$xArray = ($Setting.CentralCredentialStoreState, $Setting.TurnPasswordManagerOffState)
	If($xArray -contains "Enabled" -or $xArray -contains "Disabled")
	{
		WriteWordLine 0 2 "User Workspace\Citrix Password Manager"
		If($Setting.CentralCredentialStoreState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Central Credential Store: " $Setting.CentralCredentialStoreState
			If($Setting.CentralCredentialStoreState -eq "Enabled")
			{
				If($Setting.CentralCredentialStorePath)
				{
					WriteWordLine 0 4 "UNC path of Central Credential Store: " $Setting.CentralCredentialStorePath
				}
				Else
				{
					WriteWordLine 0 4 "No UNC path to Central Credential Store entered"
				}
			}
		}
		If($Setting.TurnPasswordManagerOffState -ne "NotConfigured")
		{
			WriteWordLine 0 3 "Do not use Citrix Password Manager: " $Setting.TurnPasswordManagerOffState
		}
	}
	
	If($Setting.StreamingDeliveryProtocolState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "User Workspace\Streamed Applications"
		WriteWordLine 0 3 "Configure delivery protocol: " $Setting.StreamingDeliveryProtocolState
		If($Setting.StreamingDeliveryProtocolState -eq "Enabled")
		{
			WriteWordLine 0 4 "Streaming Delivery Protocol option: " 
			switch ($Setting.StreamingDeliveryOption)
			{
				"Unknown"                {WriteWordLine 0 5 "Unknown"}
				"ForceServerAccess"      {WriteWordLine 0 5 "Do not allow applications to stream to the client"}
				"ForcedStreamedDelivery" {WriteWordLine 0 5 "Force applications to stream to the client"}
				Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
			}
		}
	}

	#Security
	If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "Security\Encryption\SecureICA encryption: " $Setting.SecureIcaEncriptionState
		If($Setting.SecureIcaEncriptionState -eq "Enabled")
		{
			WriteWordLine 0 3 "Encryption level: " -nonewline
			switch ($Setting.SecureIcaEncriptionLevel)
			{
				"Unknown" {WriteWordLine 0 0 "Unknown encryption"}
				"Basic"   {WriteWordLine 0 0 "Basic"}
				"LogOn"   {WriteWordLine 0 0 "RC5 (128 bit) logon only"}
				"Bits40"  {WriteWordLine 0 0 "RC5 (40 bit)"}
				"Bits56"  {WriteWordLine 0 0 "RC5 (56 bit)"}
				"Bits128" {WriteWordLine 0 0 "RC5 (128 bit)"}
				Default {WriteWordLine 0 0 "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"}
			}
		}
	}
	
	If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
	{
		WriteWordLine 0 2 "Service Level\Session Importance: " $Setting.SessionImportanceState
		If($Setting.SessionImportanceState -eq "Enabled")
		{
			WriteWordLine 0 3 "Importance level: " $Setting.SessionImportanceLevel
		}
	}
}

if (!(Check-NeededPSSnapins "Citrix.XenApp.Commands")){
    #We're missing Citrix Snapins that we need
    write-error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 5 Server? Script will now close."
    break
}

CheckWordPreReq

write-verbose "Getting Farm data"
$farm = Get-XAFarm -EA 0
If( $? )
{
	#first check to make sure this is a XenApp 5 farm
	If($Farm.ServerVersion.ToString().SubString(0,1) -ne "6")
	{
		If($Farm.ServerVersion.ToString().SubString(0,1) -eq "4")
		{
			$FarmOS = "2003"
		}
		Else
		{
			$FarmOS = "2008"
		}
		write-verbose "Farm OS is $($FarmOS)"
		#this is a XenApp 5 farm, script can proceed
		#XenApp 5 for server 2003 shows as version 4.6
		#XenApp 5 for server 2008 shows as version 5.0
	}
	Else
	{
		#this is not a XenApp 5 farm, script cannot proceed
		write-warning "This script is designed for XenApp 5 and should not be run on XenApp 6.x"
		Return 1
	}
	
	$FarmName = $farm.FarmName
	$Title="Inventory Report for the $($FarmName) Farm"
	$filename="$($pwd.path)\$($farm.FarmName).docx"
} 
Else 
{
	$FarmName = "Unable to retrieve"
	$Title="XenApp 5 Farm Inventory Report"
	$filename="$($pwd.path)\XenApp 5 Farm Inventory.docx"
	write-warning "Farm information could not be retrieved"
}
 
$farm = $null
write-verbose "Setting up Word"
#these values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
$wdSeekPrimaryFooter = 4
$wdAlignPageNumberRight = 2
$wdStory = 6
$wdMove = 0
$wdSeekMainDocument = 0
$wdColorGray15 = 14277081
$wdColorRed = 255
$wdColorBlack = 0

# Setup word for output
write-verbose "Create Word comObject.  Ignore the next message."
$Word = New-Object -comobject "Word.Application"
$WordVersion = [int] $Word.Version
If ( $WordVersion -eq 14)
{
	write-verbose "Running Microsoft Word 2010"
	$WordProduct = "Word 2010"
}
Elseif ( $WordVersion -eq 12)
{
	write-verbose "Running Microsoft Word 2007"
	$WordProduct = "Word 2007"
}
Elseif ( $WordVersion -eq 11)
{
	write-verbose "Running Microsoft Word 2003"
	Write-error "This script does not work with Word 2003. Script will end."
	$word.quit()
	exit
}
Else
{
	Write-error "You are running an untested or unsupported version of Microsoft Word.  Script will end."
	$word.quit()
	exit
}

write-verbose "Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		write-error "Company Name cannot be blank.  Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.  Script cannot continue."
		$Word.Quit()
		exit
	}
}

write-verbose "Validate cover page"
$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	write-error "For $WordProduct, $CoverPage is not a valid Cover Page option.  Script cannot continue."
	$Word.Quit()
	exit
}

Write-Verbose "Company Name: $CompanyName"
Write-Verbose "Cover Page  : $CoverPage"
Write-Verbose "User Name   : $UserName"
Write-Verbose "Farm Name   : $FarmName"
Write-Verbose "Title       : $Title"
Write-Verbose "Filename    : $filename"

$Word.Visible = $False

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
#down to $global:configlog = $false is from Jeff Hicks
write-verbose "Load Word Templates"
$CoverPagesExist = $False
$word.Templates.LoadBuildingBlocks()
If ( $WordVersion -eq 12)
{
	#word 2007
	$BuildingBlocks=$word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	#word 2010
	$BuildingBlocks=$word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

If($BuildingBlocks -ne $Null)
{
	$CoverPagesExist = $True
	$part=$BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
}
Else
{
	$CoverPagesExist = $False
}

write-verbose "Create empty word doc"
$Doc = $Word.Documents.Add()
$global:Selection = $Word.Selection

#Disable Spell and Grammer Check to resolve issue and improve performance (from Pat Coughlin)
write-verbose "disable spell checking"
$Word.Options.CheckGrammarAsYouType=$false
$Word.Options.CheckSpellingAsYouType=$false

If($CoverPagesExist)
{
	#insert new page, getting ready for table of contents
	write-verbose "insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | out-null
	$selection.InsertNewPage()

	#table of contents
	write-verbose "table of contents"
	$toc=$BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
	$toc.insert($selection.Range,$True) | out-null
}
Else
{
	write-verbose "Cover Pages are not installed."
	write-warning "Cover Pages are not installed so this report will not have a cover page."
	write-verbose "Table of Contents are not installed."
	write-warning "Table of Contents are not installed so this report will not have a Table of Contents."
}

#set the footer
write-verbose "set the footer"
[string]$footertext="Report created by $username"

#get the footer
write-verbose "get the footer and format font"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekPrimaryFooter
#get the footer and format font
$footers=$doc.Sections.Last.Footers
foreach ($footer in $footers) 
{
	if ($footer.exists) 
	{
		$footer.range.Font.name="Calibri"
		$footer.range.Font.size=8
		$footer.range.Font.Italic=$True
		$footer.range.Font.Bold=$True
	}
} #end Foreach
write-verbose "Footer text"
$selection.HeaderFooter.Range.Text=$footerText

#add page numbering
write-verbose "add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
write-verbose "return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
write-verbose "move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 

#process the nodes in the Delivery Services Console (XA5/2003) and the Access Management Console (XA5/2008)

# Get farm information
write-verbose "Getting Farm Configuration data"
$global:Server2008 = $False
$global:ConfigLog = $False
$farm = Get-XAFarmConfiguration -EA 0

If( $? )
{
	If($CoverPagesExist)
	{
		#only need the blank page inserted if there is a Table of Contents
		$selection.InsertNewPage()
	}
	WriteWordLine 1 0 "Farm Configuration Settings"
	
	WriteWordLine 2 0 "Farm-wide"

	WriteWordLine 0 1 "Connection Access Controls"
	
	switch ($Farm.ConnectionAccessControls)
	{
		"AllowAnyConnection" {WriteWordLine 0 2 "Any connections"}
		"AllowOneTypeOnly"   {
						If($FarmOS -eq "2003")
						{
							WriteWordLine 0 2 "Citrix Access Gateway, Citrix online plug-in, and Web Interface connections only"
						}
						Else
						{
							WriteWordLine 0 2 "Citrix Access Gateway, Citrix XenApp plug-in, and Web Interface connections only"
						}
					}
		"AllowMultipleTypes" {WriteWordLine 0 2 "Citrix Access Gateway connections only"}
		Default {WriteWordLine 0 0 "Connection Access Controls could not be determined: $($Farm.ConnectionAccessControls)"}
	}

	WriteWordLine 0 1 "Connection Limits" 
	WriteWordLine 0 2 "Connections per user"
	WriteWordLine 0 3 "Maximum connections per user: " -NoNewLine
	If($Farm.ConnectionLimitsMaximumPerUser -eq -1)
	{
		WriteWordLine 0 0 "No limit set"
	}
	Else
	{
		WriteWordLine 0 0 $Farm.ConnectionLimitsMaximumPerUser
	}
	If($Farm.ConnectionLimitsEnforceAdministrators)
	{
		WriteWordLine 0 3 "Enforce limit on administrators"
	}
	Else
	{
		WriteWordLine 0 3 "Do not enforce limit on administrators"
	}

	If($Farm.ConnectionLimitsLogOverLimits)
	{
		WriteWordLine 0 3 "Log over-the-limit denials"
	}
	Else
	{
		WriteWordLine 0 3 "Do not log over-the-limit denials"
	}

	#For Server 2003, the Isolation Environment section is not returned by Citrix
	
	WriteWordLine 0 1 "Health Monitoring & Recovery"
	WriteWordLine 0 2 "Limit server for load balancing"
	WriteWordLine 0 3 "Limit servers (%): " $Farm.HmrMaximumServerPercent

	WriteWordLine 0 1 "Configuration Logging"
	If($Farm.ConfigLogEnabled)
	{
		$global:ConfigLog = $True

		WriteWordLine 0 2 "Database configuration"
		WriteWordLine 0 3 "Database type: " -nonewline
		switch ($Farm.ConfigLogDatabaseType)
		{
			"SqlServer" {WriteWordLine 0 0 "Microsoft SQL Server"}
			"Oracle"    {WriteWordLine 0 0 "Oracle"}
			Default {WriteWordLine 0 0 "Database type could not be determined: $($Farm.ConfigLogDatabaseType)"}
		}
		If($Farm.ConfigLogDatabaseAuthenticationMode -eq "Native")
		{
			WriteWordLine 0 3 "Use SQL Server authentication"
		}
		Else
		{
			WriteWordLine 0 3 "Use Windows integrated security"
		}

		WriteWordLine 0 3 "Connection String: " -NoNewLine

		$StringMembers = "`n`t`t`t`t`t" + $Farm.ConfigLogDatabaseConnectionString.replace(";","`n`t`t`t`t`t")
		
		WriteWordLine 0 3 $StringMembers -NoNewLine
		WriteWordLine 0 0 "User name=" $Farm.ConfigLogDatabaseUserName

		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 3 "Log administrative tasks to logging database: " -nonewline
		}
		Else
		{
			WriteWordLine 0 3 "Log administrative tasks to Configuration Logging database: " -nonewline
		}
		If($Farm.ConfigLogEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 3 "Allow changes to the farm when database is disconnected: " -nonewline
		}
		Else
		{
			WriteWordLine 0 3 "Allow changes to the farm when logging database is disconnected: " -nonewline
		}
		
		If($Farm.ConfigLogChangesWhileDisconnectedAllowed)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Require admins to enter database credentials before clearing the log: " -nonewline
		If($Farm.ConfigLogCredentialsOnClearLogRequired)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
	}
	Else
	{
		WriteWordLine 0 2 "Configuration logging is not enabled"
	}
		
	WriteWordLine 0 1 "Memory/CPU"

	WriteWordLine 0 2 "Applications that memory optimization ignores: "
	If($Farm.MemoryOptimizationExcludedApplications)
	{
		ForEach($App in $Farm.MemoryOptimizationExcludedApplications)
		{
			WriteWordLine 0 3 $App
		}
	}
	Else
	{
		WriteWordLine 0 3 "No applications are listed"
	}

	WriteWordLine 0 2 "Optimization interval: " $Farm.MemoryOptimizationScheduleType

	If($Farm.MemoryOptimizationScheduleType -eq "Weekly")
	{
		WriteWordLine 0 2 "Day of week: " $Farm.MemoryOptimizationScheduleDayOfWeek
	}
	If($Farm.MemoryOptimizationScheduleType -eq "Monthly")
	{
		WriteWordLine 0 2 "Day of month: " $Farm.MemoryOptimizationScheduleDayOfMonth
	}

	WriteWordLine 0 2 "Optimization time: " $Farm.MemoryOptimizationScheduleTime
	WriteWordLine 0 2 "Memory optimization user: " -nonewline
	If($Farm.MemoryOptimizationLocalSystemAccountUsed)
	{
		WriteWordLine 0 0 "Use local system account"
	}
	Else
	{
		WriteWordLine 0 0 $Farm.MemoryOptimizationUser
	}
	
	WriteWordLine 0 1 "XenApp"
	WriteWordLine 0 2 "General"
	WriteWordLine 0 3 "Respond to client broadcast messages"
	WriteWordLine 0 4 "Data collectors: " -nonewline
	If($Farm.RespondDataCollectors)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 4 "RAS servers: " -nonewline
	If($Farm.RespondRasServers)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 3 "Client time zones"
	WriteWordLine 0 4 "Use client's local time: " -nonewline
	If($Farm.ClientLocalTimeEnabled)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 4 "Estimate local time for clients: " -nonewline
	If($Farm.ClientLocalTimeEstimationEnabled)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 3 "XML Service DNS address resolution: " -nonewline
	If($Farm.DNSAddressResolution)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 3 "Novell Directory Services"
	WriteWordLine 0 4 "NDS preferred tree: " -NoNewLine
	If($Farm.NdsPreferredTree)
	{
		WriteWordLine 0 0 $Farm.NdsPreferredTree
	}
	Else
	{
		WriteWordLine 0 0 "No NDS Tree entered"
	}
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 3 "Enable 32 bit icon color depth: " -nonewline
	}
	Else
	{
		WriteWordLine 0 3 "Enhanced icon support: " -nonewline
	}
	If($Farm.EnhancedIconEnabled)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}

	WriteWordLine 0 2 "Shadow Policies"
	WriteWordLine 0 3 "Merge shadowers in multiple policies: " -nonewline
	If($Farm.ShadowPoliciesMerge)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}

	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 1 "HDX Broadcast"
		WriteWordLine 0 2 "Session Reliability"
	}
	Else
	{
		WriteWordLine 0 1 "Session Reliability"
	}
	
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 3 "Allow users to view sessions during broken connection: " -nonewline
	}
	Else
	{
		WriteWordLine 0 2 "Keep sessions open during loss of network connectivity: " -nonewline
	}
	If($Farm.SessionReliabilityEnabled)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 3 "Port number (default 2598): " $Farm.SessionReliabilityPort
	}
	Else
	{
		WriteWordLine 0 2 "Port number (default 2598): " $Farm.SessionReliabilityPort
	}
	
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 3 "Seconds to keep sessions active: " $Farm.SessionReliabilityTimeout
	}
	Else
	{
		WriteWordLine 0 2 "Seconds to keep sessions open: " $Farm.SessionReliabilityTimeout
	}

	WriteWordLine 0 1 "Citrix Streaming Server"
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 2 "Log application events to event log: " -nonewline
	}
	Else
	{
		WriteWordLine 0 2 "Log application streaming events to event log: " -nonewline
	}
	If($Farm.StreamingLogEvents)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 2 "Trust Citrix Delivery Clients: " -nonewline
	}
	Else
	{
		WriteWordLine 0 2 "Trust XenApp Plugin for Streamed Apps: " -nonewline
	}
	If($Farm.StreamingTrustCLient)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}

	If($FarmOS -eq "2008")
	{
		WriteWordLine 0 1 "Restart Options"
		WriteWordLine 0 2 "Message Options"
		WriteWordLine 0 3 "Send message to logged-on users before server restart: " -nonewline
		If($Farm.RestartSendMessage)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 3 "Send first message before restart: $($Farm.RestartMessageWait) minutes"
		WriteWordLine 0 3 "Send reminder message every: $($Farm.RestartMessageInterval) minutes"
		If($Farm.RestartCustomMessageEnabled)
		{
			WriteWordLine 0 3 "Additional text for restart message:"
			WriteWordLine 0 3 $Farm.RestartCustomMessage
		}
		If($Farm.RestartDisabledLogOnsInterval -gt 0)
		{
			WriteWordLine 0 3 "Disable logons before restart: " -nonewline
			WriteWordLine 0 0 "Yes"
			WriteWordLine 0 3 "Before restart, logons disabled by: $($Farm.RestartDisabledLogOnsInterval) minutes"
		}
		Else
		{
			WriteWordLine 0 3 "Disable logons before restart: No"
		}
	}
	
	WriteWordLine 0 1 "Virtual IP"
	WriteWordLine 0 2 "Address Configuration"
	WriteWordLine 0 3 "Virtual IP address ranges:"

	$VirtualIPs = Get-XAVirtualIPRange -EA 0
	If($? -and $VirtualIPs)
	{
		ForEach($VirtualIP in $VirtualIPs)
		{
			WriteWordLine 0 4 "IP Range: " $VirtualIP
		}
	}
	Else
	{
		WriteWordLine 0 4 "No virtual IP address range defined"
	}
	$VirtualIPs = $Null

	WriteWordLine 0 3 "Enable logging of IP address assignment and release: " -nonewline
	If($Farm.VirtualIPLoggingEnabled)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 2 "Process Configuration"
	WriteWordLine 0 3 "Virtual IP Processes"
	If($Farm.VirtualIPProcesses)
	{
		WriteWordLine 0 4 "Monitor the following processes:"
		ForEach($Process in $Farm.VirtualIPProcesses)
		{
			WriteWordLine 0 5 "Process: " $Process
		}
	}
	Else
	{
		WriteWordLine 0 4 "No virtual IP processes defined"
	}
	WriteWordLine 0 3 "Virtual Loopback Processes"
	If($Farm.VirtualIPLoopbackProcesses)
	{
		WriteWordLine 0 4 "Monitor the following processes:"
		ForEach($Process in $Farm.VirtualIPLoopbackProcesses)
		{
			WriteWordLine 0 5 "Process: " $Process
		}
	}
	Else
	{
		WriteWordLine 0 4 "No virtual IP Loopback processes defined"
	}
		
	$selection.InsertNewPage()
	WriteWordLine 2 0 "Server Default"
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 1 "HDX Broadcast"
	}
	Else
	{
		WriteWordLine 0 1 "ICA"
	}
	WriteWordLine 0 2 "Auto Client Reconnect"
	If($Farm.AcrEnabled)
	{
		WriteWordLine 0 3 "Reconnect automatically"
		WriteWordLine 0 3 "Log automatic reconnection attempts: " -NoNewLine

		If($Farm.AcrLogReconnections)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
	}
	Else
	{
		WriteWordLine 0 4 "Require user authentication"
	}
	
	WriteWordLine 0 2 "Display"
	WriteWordLine 0 3 "Discard queued image that is replaced by another image: " -nonewline
	If($Farm.DisplayDiscardQueuedImages)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 3 "Cache image to make scrolling smoother: " -nonewline
	If($Farm.DisplayCacheImageForSmoothScrolling)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}
	WriteWordLine 0 3 "Maximum memory to use for each session's graphics (KB): " $Farm.DisplayMaximumGraphicsMemory
	WriteWordLine 0 3 "Degradation bias"
	If($Farm.DisplayDegradationBias -eq "Resolution")
	{
		WriteWordLine 0 4 "Degrade resolution first"
	}
	Else
	{
		WriteWordLine 0 4 "Degrade color depth first"
	}
	WriteWordLine 0 3 "Notify user of session degradation: " -nonewline
	If($Farm.DisplayNotifyUser)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}

	WriteWordLine 0 2 "Keep-Alive"
	If($Farm.KeepAliveEnabled)
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 3 "HDX Broadcast Keep-Alive time-out value (seconds): " $Farm.KeepAliveTimeout
		}
		Else
		{
			WriteWordLine 0 3 "ICA Keep-Alive time-out value (seconds): " $Farm.KeepAliveTimeout
		}
	}
	Else
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 3 "HDX Broadcast Keep-Alive is not enabled"
		}
		Else
		{
			WriteWordLine 0 3 "ICA Keep-Alive is not enabled"
		}
	}
	
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 2 "Remote Console Connections"
		WriteWordLine 0 3 "Remote connections to the console: " -nonewline
		If($Farm.RemoteConsoleEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
	}
	
	#For server 2003, Isolation Environment data is not available
	
	WriteWordLine 0 1 "License Server"
	WriteWordLine 0 2 "Name: " $Farm.LicenseServerName
	WriteWordLine 0 2 "Port number (default 27000): " $Farm.LicenseServerPortNumber
	
	WriteWordLine 0 1 "Memory/CPU"
	WriteWordLine 0 2 "CPU Utilization Management: " -NoNewLine
	If($Farm.CpuManagementLevel.ToString() -eq "255")
	{
		WriteWordLine 0 0 "Cannot be determined for XenApp 5 on Server 2003"
	}
	Else
	{
		WriteWordLine 0 0 "" -nonewline
		switch ($Farm.CpuManagementLevel)
		{
			"NoManagement"  {WriteWordLine 0 0 "No CPU utilization management"}
			"Fair"          {WriteWordLine 0 0 "Fair sharing of CPU between sessions"}
			"ResourceBased" {WriteWordLine 0 0 "CPU Sharing based on Resource Allotments"}
			Default {WriteWordLine 0 0 "CPU Utilization Management could not be determined: $($Farm.CpuManagementLevel)"}
		}
	}
	WriteWordLine 0 2 "Memory Optimization: " -nonewline
	If($Farm.MemoryOptimizationEnabled)
	{
		WriteWordLine 0 0 "Enabled"
	}
	Else
	{
		WriteWordLine 0 0 "Not Enabled"
	}
	
	WriteWordLine 0 1 "Health Monitoring & Recovery"
	If($Farm.HmrEnabled)
	{
		$HmrTests = Get-XAHmrTest -EA 0 | Sort-Object TestName
		If($?)
		{
			ForEach($HmrTest in $HmrTests)
			{
				WriteWordLine 0 2 "Test Name`t: " $Hmrtest.TestName
				WriteWordLine 0 2 "Interval`t`t: " $Hmrtest.Interval
				WriteWordLine 0 2 "Threshold`t: " $Hmrtest.Threshold
				WriteWordLine 0 2 "Time-out`t: " $Hmrtest.Timeout
				WriteWordLine 0 2 "Test File Name`t: " $Hmrtest.FilePath
				If(![String]::IsNullOrEmpty($Hmrtest.Arguments))
				{
					WriteWordLine 0 2 "Arguments`t: " $Hmrtest.Arguments
				}
				WriteWordLine 0 2 "Recovery Action : " -nonewline
				switch ($Hmrtest.RecoveryAction)
				{
					"AlertOnly"                     {WriteWordLine 0 0 "Alert Only"}
					"RemoveServerFromLoadBalancing" {WriteWordLine 0 0 "Remove Server from load balancing"}
					"RestartIma"                    {WriteWordLine 0 0 "Restart IMA"}
					"ShutdownIma"                   {WriteWordLine 0 0 "Shutdown IMA"}
					"RebootServer"                  {WriteWordLine 0 0 "Reboot Server"}
					Default {WriteWordLine 0 0 "Recovery Action could not be determined: $($Hmrtest.RecoveryAction)"}
				}
				WriteWordLine 0 0 ""
			}
		}
		Else
		{
			WriteWordLine 0 2 "Health Monitoring & Recovery Tests could not be retrieved"
		}
	}
	Else
	{
		WriteWordLine 0 2 "Health Monitoring & Recovery is not enabled"
	}

	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 1 "HDX Plug and Play"
	}
	Else
	{
		WriteWordLine 0 1 "XenApp"
	}
	WriteWordLine 0 2 "Content redirection from server to client: " -nonewline
	If($Farm.ContentRedirectionEnabled)
	{
		WriteWordLine 0 0 "Yes"
	}
	Else
	{
		WriteWordLine 0 0 "No"
	}

	If($FarmOS -eq "2008")
	{
		WriteWordLine 0 2 "Remote Console Connections"
		WriteWordLine 0 3 "Remote connections to the console: " -nonewline
		If($Farm.RemoteConsoleEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
	}

	WriteWordLine 0 1 "SNMP"
	If($Farm.SnmpEnabled)
	{
		WriteWordLine 0 2 "Send session traps to selected SNMP agent on all farm servers"
		WriteWordLine 0 3 "SNMP agent session traps"
		WriteWordLine 0 4 "Logon`t`t`t: " -nonewline
		If($Farm.SnmpLogonEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 4 "Logoff`t`t`t: " -nonewline
		If($Farm.SnmpLogoffEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 4 "Disconnect`t`t: " -nonewline
		If($Farm.SnmpDisconnectEnabled)
		{
			WriteWordLine 0 0 "Yes"
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
		WriteWordLine 0 4 "Session limit per server`t: " -nonewline
		If($Farm.SnmpLimitEnabled)
		{
			WriteWordLine 0 0 " " $Farm.SnmpLimitPerServer
		}
		Else
		{
			WriteWordLine 0 0 "No"
		}
	}
	Else
	{
		WriteWordLine 0 2 "SNMP is not enabled"
	}

	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 1 "HDX 3D"
	}
	Else
	{
		WriteWordLine 0 1 "SpeedScreen"
	}
	If($Farm.BrowserAccelerationEnabled)
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 2 "HDX 3D Browser Acceleration is enabled"
		}
		Else
		{
			WriteWordLine 0 2 "SpeedScreen Browser Acceleration is enabled"
		}
		If($Farm.BrowserAccelerationCompressionEnabled)
		{
			WriteWordLine 0 3 "Compress JPEG images to improve bandwidth"
			WriteWordLine 0 4 "Image compression levels: " $Farm.BrowserAccelerationCompressionLevel
			If($Farm.BrowserAccelerationVariableImageCompression)
			{
				WriteWordLine 0 4 "Adjust compression level based on available bandwidth"
			}
			Else
			{
				WriteWordLine 0 4 "Do not adjust compression level based on available bandwidth"
			}
		}
		Else
		{
			WriteWordLine 0 3 "Do not compress JPEG images to improve bandwidth"
		}
	}
	Else
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 2 "HDX 3D Browser Acceleration is disabled"
		}
		Else
		{
			WriteWordLine 0 2 "SpeedScreen Browser Acceleration is disabled"
		}
	}
	
	If($FarmOS -eq "2003")
	{
		WriteWordLine 0 1 "HDX Mediastream"
	}
	If($Farm.FlashAccelerationEnabled)
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 2 "Enable Flash for XenApp sessions"
			WriteWordLine 0 3 "Server-side acceleration: " -nonewline
			switch ($Farm.FlashAccelerationOption)
			{
				"AllConnections" {WriteWordLine 0 0 "Accelerate for restricted bandwidth connections"}
				"Unknown"        {WriteWordLine 0 0 "Do not accelerate"}
				"NoOptimization" {WriteWordLine 0 0 "Accelerate for all connections"}
				Default {WriteWordLine 0 0 "Server-side acceleration could not be determined: $($Farm.FlashAccelerationOption)"}
			}
		}
		Else
		{
			WriteWordLine 0 2 "Enable Adobe Flash Player"
			switch ($Farm.FlashAccelerationOption)
			{
				"AllConnections" {WriteWordLine 0 3 "Accelerate for restricted bandwidth connections"}
				"Unknown"        {WriteWordLine 0 3 "Do not accelerate"}
				"NoOptimization" {WriteWordLine 0 3 "Accelerate for all connections"}
				Default {WriteWordLine 0 0 "Server-side acceleration could not be determined: $($Farm.FlashAccelerationOption)"}
			}
		}
		
	}
	Else
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 2 "Flash is not enabled for XenApp sessions"
		}
		Else
		{
			WriteWordLine 0 2 "Adobe Flash is not enabled"
		}
	}
	If($Farm.MultimediaAccelerationEnabled)
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 2 "Multimedia Acceleration is enabled"
			WriteWordLine 0 3 "Multimedia Acceleration (Network buffering)"
		}
		Else
		{
			WriteWordLine 0 2 "SpeedScreen Multimedia Acceleration is enabled"
		}
		If($Farm.MultimediaAccelerationDefaultBuffer)
		{
			WriteWordLine 0 3 "Use the default buffer of 5 seconds"
		}
		Else
		{
			WriteWordLine 0 3 "Custom buffer in seconds: " $Farm.MultimediaAccelerationCustomBuffer
		}
	}
	Else
	{
		If($FarmOS -eq "2003")
		{
			WriteWordLine 0 2 "Multimedia Acceleration is disabled"
		}
		Else
		{
			WriteWordLine 0 2 "SpeedScreen Multimedia Acceleration is disabled"
		}
	}
	
	WriteWordLine 0 0 "Offline Access"
	WriteWordLine 0 1 "Users"
	If($Farm.OfflineAccounts)
	{
		WriteWordLine 0 2 "Configured users:"
		ForEach($User in $Farm.OfflineAccounts)
		{
			WriteWordLine 0 3 $User
		}
	}
	Else
	{
		WriteWordLine 0 2 "No users configured"
	}

	WriteWordLine 0 1 "Offline License Settings"
	WriteWordLine 0 2 "License period days: " $Farm.OfflineLicensePeriod

} 
Else 
{
	Write-warning "Farm information could not be retrieved"
}
$farm = $null

write-verbose "Processing Administrators"
$Administrators = Get-XAAdministrator -EA 0 | sort-object AdministratorName

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Administrators:"
	ForEach($Administrator in $Administrators)
	{
		WriteWordLine 2 0 $Administrator.AdministratorName
		WriteWordLine 0 1 "Administrator type: "$Administrator.AdministratorType -nonewline
		WriteWordLine 0 0 " Administrator"
		WriteWordLine 0 1 "Administrator account is " -NoNewLine
		If($Administrator.Enabled)
		{
			WriteWordLine 0 0 "Enabled" 
		} 
		Else
		{
			WriteWordLine 0 0 "Disabled" 
		}
		If(!([String]::IsNullOrEmpty($Administrator.EmailAddress) -and [String]::IsNullOrEmpty($Administrator.SmsNumber) -and [String]::IsNullOrEmpty($Administrator.SmsGateway)))
		{
			WriteWordLine 0 1 "Alert Contact Details"
			WriteWordLine 0 2 "E-mail`t`t: " $Administrator.EmailAddress
			WriteWordLine 0 2 "SMS Number`t: " $Administrator.SmsNumber
			WriteWordLine 0 2 "SMS Gateway`t: " $Administrator.SmsGateway
		}
		If ($Administrator.AdministratorType -eq "Custom") 
		{
			WriteWordLine 0 1 "Farm Privileges:"
			ForEach($farmprivilege in $Administrator.FarmPrivileges) 
			{
				#WriteWordLine 0 2 $farmprivilege
				switch ($farmprivilege)
				{
					"Unknown"                   {WriteWordLine 0 2 "Unknown"}
					"ViewFarm"                  {WriteWordLine 0 2 "View farm management"}
					"EditZone"                  {WriteWordLine 0 2 "Edit zones"}
					"EditConfigurationLog"      {WriteWordLine 0 2 "Configure logging for the farm"}
					"EditFarmOther"             {WriteWordLine 0 2 "Edit all other farm settings"}
					"ViewAdmins"                {WriteWordLine 0 2 "View Citrix administrators"}
					"LogOnConsole"              {WriteWordLine 0 2 "Log on to console"}
					"LogOnWIConsole"            {WriteWordLine 0 2 "Logon on to Web Interface console"}
					"ViewLoadEvaluators"        {WriteWordLine 0 2 "View load evaluators"}
					"AssignLoadEvaluators"      {WriteWordLine 0 2 "Assign load evaluators"}
					"EditLoadEvaluators"        {WriteWordLine 0 2 "Edit load evaluators"}
					"ViewLoadBalancingPolicies" {WriteWordLine 0 2 "View load balancing policies"}
					"EditLoadBalancingPolicies" {WriteWordLine 0 2 "Edit load balancing policies"}
					"ViewPrinterDrivers"        {WriteWordLine 0 2 "View printer drivers"}
					"ReplicatePrinterDrivers"   {WriteWordLine 0 2 "Replicate printer drivers"}
					"EditUserPolicies"          {WriteWordLine 0 2 "Edit User Policies"}
					"ViewUserPolicies"          {WriteWordLine 0 2 "View User Policies"}
					"EditOtherPrinterSettings"  {WriteWordLine 0 2 "Edit All Other Printer Settings"}
					"EditPrinterDrivers"        {WriteWordLine 0 2 "Edit Printer Drivers"}
					"EditPrinters"              {WriteWordLine 0 2 "Edit Printers"}
					"ViewPrintersAndDrivers"    {WriteWordLine 0 2 "View Printers and Printer Drovers"}
					Default {WriteWordLine 0 2 "Farm privileges could not be determined: $($farmprivilege)"}
				}
			}
	
			WriteWordLine 0 1 "Folder Privileges:"
			ForEach($folderprivilege in $Administrator.FolderPrivileges) 
			{
				#$test = $folderprivilege.ToString()
				#$folderlabel = $test.substring(0, $test.IndexOf(":") + 1)
				#WriteWordLine 0 2 $folderlabel
				#$test1 = $test.substring($test.IndexOf(":") + 1)
				#$folderpermissions = $test1.replace(",","`n`t`t`t")
				#WriteWordLine 0 3 $folderpermissions
				WriteWordLine 0 2 $FolderPrivilege.FolderPath
				ForEach($FolderPermission in $FolderPrivilege.FolderPrivileges)
				{
					switch ($folderpermission)
					{
						"Unknown"                          {WriteWordLine 0 3 "Unknown"}
						"ViewApplications"                 {WriteWordLine 0 3 "View applications"}
						"EditApplications"                 {WriteWordLine 0 3 "Edit applications"}
						"TerminateProcessApplication"      {WriteWordLine 0 3 "Terminate process that is created as a result of launching a published application"}
						"AssignApplicationsToServers"      {WriteWordLine 0 3 "Assign applications to servers"}
						"ViewServers"                      {WriteWordLine 0 3 "View servers"}
						"EditOtherServerSettings"          {WriteWordLine 0 3 "Edit other server settings"}
						"RemoveServer"                     {WriteWordLine 0 3 "Remove a bad server from farm"}
						"TerminateProcess"                 {WriteWordLine 0 3 "Terminate processes on a server"}
						"ViewSessions"                     {WriteWordLine 0 3 "View ICA/RDP sessions"}
						"ConnectSessions"                  {WriteWordLine 0 3 "Connect sessions"}
						"DisconnectSessions"               {WriteWordLine 0 3 "Disconnect sessions"}
						"LogOffSessions"                   {WriteWordLine 0 3 "Log off sessions"}
						"ResetSessions"                    {WriteWordLine 0 3 "Reset sessions"}
						"SendMessages"                     {WriteWordLine 0 3 "Send messages to sessions"}
						"ViewWorkerGroups"                 {WriteWordLine 0 3 "View worker groups"}
						"AssignApplicationsToWorkerGroups" {WriteWordLine 0 3 "Assign applications to worker groups"}
						"AssignApplications"               {WriteWordLine 0 3 "Assign Application to Servers"}
						"EditServerSnmpSettings"           {WriteWordLine 0 3 "Edit SNMP Settings"}
						"EditLicenseServer"                {WriteWordLine 0 3 "Edit License Server Settings"}
						Default {WriteWordLine 0 3 "Folder permission could not be determined: $($folderpermission)"}
					}
				}
			}
		}		
		WriteWordLine 0 0 " "
	}
}
Else 
{
	Write-warning "Administrator information could not be retrieved"
}

$Administrators = $null

write-verbose "Processing Application"
$Applications = Get-XAApplication -EA 0 | sort-object FolderPath, DisplayName

If( $? -and $Applications)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Applications:"
	ForEach($Application in $Applications)
	{
		$AppServerInfoResults = $False
		$AppServerInfo = Get-XAApplicationReport -BrowserName $Application.BrowserName -EA 0
		If( $? )
		{
			$AppServerInfoResults = $True
		}
		$streamedapp = $False
		If($Application.ApplicationType -Contains "streamedtoclient" -or $Application.ApplicationType -Contains "streamedtoserver")
		{
			$streamedapp = $True
		}
		#name properties
		WriteWordLine 2 0 $Application.DisplayName
		WriteWordLine 0 1 "Application name`t`t: " $Application.BrowserName
		WriteWordLine 0 1 "Disable application`t`t: " -NoNewLine
		#weird, if application is enabled, it is disabled!
		If ($Application.Enabled) 
		{
			WriteWordLine 0 0 "No"
		} 
		Else
		{
			WriteWordLine 0 0 "Yes"
			WriteWordLine 0 1 "Hide disabled application`t: " -nonewline
			If($Application.HideWhenDisabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}

		If(![String]::IsNullOrEmpty( $Application.Description))
		{
			WriteWordLine 0 1 "Application description`t`t: " $Application.Description
		}
		
		#type properties
		WriteWordLine 0 1 "Application Type`t`t: " -nonewline
		switch ($Application.ApplicationType)
		{
			"Unknown"                            {WriteWordLine 0 0 "Unknown"}
			"ServerInstalled"                    {WriteWordLine 0 0 "Installed application"}
			"ServerDesktop"                      {WriteWordLine 0 0 "Server desktop"}
			"Content"                            {WriteWordLine 0 0 "Content"}
			"StreamedToServer"                   {WriteWordLine 0 0 "Streamed to server"}
			"StreamedToClient"                   {WriteWordLine 0 0 "Streamed to client"}
			"StreamedToClientOrInstalled"        {WriteWordLine 0 0 "Streamed if possible, otherwise accessed from server as Installed application"}
			"StreamedToClientOrStreamedToServer" {WriteWordLine 0 0 "Streamed if possible, otherwise Streamed to server"}
			Default {WriteWordLine 0 0 "Application Type could not be determined: $($Application.ApplicationType)"}
		}
		If(![String]::IsNullOrEmpty( $Application.FolderPath))
		{
			WriteWordLine 0 1 "Folder path`t`t`t: " $Application.FolderPath
		}
		If(![String]::IsNullOrEmpty( $Application.ContentAddress))
		{
			WriteWordLine 0 1 "Content Address`t`t: " $Application.ContentAddress
		}
	
		#if a streamed app
		If($streamedapp)
		{
			WriteWordLine 0 1 "Citrix streaming app profile address`t`t: " 
			WriteWordLine 0 2 $Application.ProfileLocation
			WriteWordLine 0 1 "App to launch from Citrix stream app profile`t: " 
			WriteWordLine 0 2 $Application.ProfileProgramName
			If(![String]::IsNullOrEmpty( $Application.ProfileProgramArguments))
			{
				WriteWordLine 0 1 "Extra command line parameters`t`t`t: " 
				WriteWordLine 0 2 $Application.ProfileProgramArguments
			}
			#if streamed, Offline access properties
			If($Application.OfflineAccessAllowed)
			{
				WriteWordLine 0 1 "Enable offline access`t`t`t`t: " -nonewline
				If($Application.OfflineAccessAllowed)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
			}
			If($Application.CachingOption)
			{
				WriteWordLine 0 1 "Cache preference`t`t`t`t: " -nonewline
				switch ($Application.CachingOption)
				{
					"Unknown"   {WriteWordLine 0 0 "Unknown"}
					"PreLaunch" {WriteWordLine 0 0 "Cache application prior to launching"}
					"AtLaunch"  {WriteWordLine 0 0 "Cache application during launch"}
					Default {WriteWordLine 0 0 "Could not be determined: $($Application.CachingOption)"}
				}
			}
		}
		
		#location properties
		If(!$streamedapp)
		{
			If(![String]::IsNullOrEmpty( $Application.CommandLineExecutable))
			{
				If($Application.CommandLineExecutable.Length -lt 40)
				{
					WriteWordLine 0 1 "Command Line`t`t`t: " $Application.CommandLineExecutable
				}
				Else
				{
					WriteWordLine 0 1 "Command Line: " 
					WriteWordLine 0 2 $Application.CommandLineExecutable
				}
			}
			If(![String]::IsNullOrEmpty( $Application.WorkingDirectory))
			{
				If($Application.WorkingDirectory.Length -lt 40)
				{
					WriteWordLine 0 1 "Working directory`t`t: " $Application.WorkingDirectory
				}
				Else
				{
					WriteWordLine 0 1 "Working directory: " 
					WriteWordLine 0 2 $Application.WorkingDirectory
				}
			}
			
			#servers properties
			If($AppServerInfoResults)
			{
				If(![String]::IsNullOrEmpty( $AppServerInfo.ServerNames))
				{
					WriteWordLine 0 1 "Servers:"
					ForEach($servername in $AppServerInfo.ServerNames)
					{
						WriteWordLine 0 2 $servername
					}
				}
			}
			Else
			{
				WriteWordLine 0 2 "Unable to retrieve a list of Servers for this application"
			}
		}
	
		#users properties
		If($Application.AnonymousConnectionsAllowed)
		{
			WriteWordLine 0 1 "Allow anonymous users: " $Application.AnonymousConnectionsAllowed
		}
		Else
		{
			If($AppServerInfoResults)
			{
				WriteWordLine 0 1 "Users:"
				ForEach($user in $AppServerInfo.Accounts)
				{
					WriteWordLine 0 2 $user
				}
			}
			Else
			{
				WriteWordLine 0 2 "Unable to retrieve a list of Users for this application"
			}
		}
	
		#shortcut presentation properties
		#application icon is ignored
		If(![String]::IsNullOrEmpty($Application.ClientFolder))
		{
			If($Application.ClientFolder.Length -lt 30)
			{
				WriteWordLine 0 1 "Client application folder`t`t`t`t: " $Application.ClientFolder
			}
			Else
			{
				WriteWordLine 0 1 "Client application folder`t`t`t`t: " 
				WriteWordLine 0 2 $Application.ClientFolder
			}
		}
		If($Application.AddToClientStartMenu)
		{
			WriteWordLine 0 1 "Add to client's start menu"
			If($Application.StartMenuFolder)
			{
				WriteWordLine 0 2 "Start menu folder`t`t`t: " $Application.StartMenuFolder
			}
		}
		If($Application.AddToClientDesktop)
		{
			WriteWordLine 0 1 "Add shortcut to the client's desktop"
		}
	
		#access control properties
		If($Application.ConnectionsThroughAccessGatewayAllowed)
		{
			WriteWordLine 0 1 "Allow connections made through AGAE`t`t: " -nonewline
			If($Application.ConnectionsThroughAccessGatewayAllowed)
			{
				WriteWordLine 0 0 "Yes"
			} 
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		If($Application.OtherConnectionsAllowed)
		{
			WriteWordLine 0 1 "Any connection`t`t`t`t`t: " -nonewline
			If($Application.OtherConnectionsAllowed)
			{
				WriteWordLine 0 0 "Yes"
			} 
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		If($Application.AccessSessionConditionsEnabled)
		{
			WriteWordLine 0 1 "Any connection that meets any of the following filters: " $Application.AccessSessionConditionsEnabled
			WriteWordLine 0 1 "Access Gateway Filters:"
			ForEach($filter in $Application.AccessSessionConditions)
			{
				WriteWordLine 0 2 $filter
			}
		}
	
		#content redirection properties
		If($AppServerInfoResults)
		{
			If($AppServerInfo.FileTypes)
			{
				WriteWordLine 0 1 "File type associations:"
				ForEach($filetype in $AppServerInfo.FileTypes)
				{
					WriteWordLine 0 3 $filetype
				}
			}
			Else
			{
				WriteWordLine 0 1 "File Type Associations for this application`t: None"
			}
		}
		Else
		{
			WriteWordLine 0 1 "Unable to retrieve the list of FTAs for this application"
		}
	
		#if streamed app, Alternate profiles
		If($streamedapp)
		{
			If($Application.AlternateProfiles)
			{
				WriteWordLine 0 1 "Primary application profile location`t`t: " $Application.AlternateProfiles
			}
		
			#if streamed app, User privileges properties
			If($Application.RunAsLeastPrivilegedUser)
			{
				WriteWordLine 0 1 "Run application as a least-privileged user account`t: " $Application.RunAsLeastPrivilegedUser
			}
		}
	
		#limits properties
		WriteWordLine 0 1 "Limit instances allowed to run in server farm`t: " -NoNewLine

		If($Application.InstanceLimit -eq -1)
		{
			WriteWordLine 0 0 "No limit set"
		}
		Else
		{
			WriteWordLine 0 0 $Application.InstanceLimit
		}
	
		WriteWordLine 0 1 "Allow only 1 instance of app for each user`t: " -NoNewLine
	
		If ($Application.MultipleInstancesPerUserAllowed) 
		{
			WriteWordLine 0 0 "No"
		} 
		Else
		{
			WriteWordLine 0 0 "Yes"
		}
	
		If($Application.CpuPriorityLevel)
		{
			WriteWordLine 0 1 "Application importance`t`t`t`t: " -nonewline
			switch ($Application.CpuPriorityLevel)
			{
				"Unknown"     {WriteWordLine 0 0 "Unknown"}
				"BelowNormal" {WriteWordLine 0 0 "Below Normal"}
				"Low"         {WriteWordLine 0 0 "Low"}
				"Normal"      {WriteWordLine 0 0 "Normal"}
				"AboveNormal" {WriteWordLine 0 0 "Above Normal"}
				"High"        {WriteWordLine 0 0 "High"}
				Default {WriteWordLine 0 0 "Application importance could not be determined: $($Application.CpuPriorityLevel)"}
			}
		}
		
		#client options properties
		WriteWordLine 0 1 "Enable legacy audio`t`t`t`t: " -nonewline
		switch ($Application.AudioType)
		{
			"Unknown" {WriteWordLine 0 0 "Unknown"}
			"None"    {WriteWordLine 0 0 "Not Enabled"}
			"Basic"   {WriteWordLine 0 0 "Enabled"}
			Default {WriteWordLine 0 0 "Enable legacy audio could not be determined: $($Application.AudioType)"}
		}
		WriteWordLine 0 1 "Minimum requirement`t`t`t`t: " -nonewline
		If($Application.AudioRequired)
		{
			WriteWordLine 0 0 "Enabled"
		}
		Else
		{
			WriteWordLine 0 0 "Disabled"
		}
		If($Application.SslConnectionEnabled)
		{
			WriteWordLine 0 1 "Enable SSL and TLS protocols`t`t`t: " -nonewline
			If($Application.SslConnectionEnabled)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
		If($Application.EncryptionLevel)
		{
			WriteWordLine 0 1 "Encryption`t`t`t`t`t: " -nonewline
			switch ($Application.EncryptionLevel)
			{
				"Unknown" {WriteWordLine 0 0 "Unknown"}
				"Basic"   {WriteWordLine 0 0 "Basic"}
				"LogOn"   {WriteWordLine 0 0 "128-Bit Login Only (RC-5)"}
				"Bits40"  {WriteWordLine 0 0 "40-Bit (RC-5)"}
				"Bits56"  {WriteWordLine 0 0 "56-Bit (RC-5)"}
				"Bits128" {WriteWordLine 0 0 "128-Bit (RC-5)"}
				Default {WriteWordLine 0 0 "Encryption could not be determined: $($Application.EncryptionLevel)"}
			}
		}
		If($Application.EncryptionRequired)
		{
			WriteWordLine 0 1 "Minimum requirement`t`t`t`t: " -nonewline
			If($Application.EncryptionRequired)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
	
		WriteWordLine 0 1 "Start app w/o waiting for printer creation`t: " -NoNewLine
		#another weird one, if True then this is Disabled
		If ($Application.WaitOnPrinterCreation) 
		{
			WriteWordLine 0 0 "No"
		} 
		Else
		{
			WriteWordLine 0 0 "Yes"
		}
		
		#appearance properties
		If($Application.WindowType)
		{
			WriteWordLine 0 1 "Session window size`t`t`t`t: " $Application.WindowType
		}
		If($Application.ColorDepth)
		{
			WriteWordLine 0 1 "Maximum color quality`t`t`t`t: " -nonewline
			switch ($Application.ColorDepth)
			{
				"Colors16"  {WriteWordLine 0 0 "16 colors"}
				"Colors256" {WriteWordLine 0 0 "256 colors"}
				"HighColor" {WriteWordLine 0 0 "High Color (16-bit)"}
				"TrueColor" {WriteWordLine 0 0 "True Color (24-bit)"}
				Default {WriteWordLine 0 0 "Maximum color quality could not be determined: $($Application.ColorDepth)"}
			}
		}
		If($Application.TitleBarHidden)
		{
			WriteWordLine 0 1 "Hide application title bar`t`t`t: " -nonewline
			If($Application.TitleBarHidden)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
		If($Application.MaximizedOnStartup)
		{
			WriteWordLine 0 1 "Maximize application at startup`t`t`t: " -nonewline
			If($Application.MaximizedOnStartup)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
	}
}
Else 
{
	Write-warning "Application information could not be retrieved"
}

$Applications = $null

#servers
write-verbose "Processing Servers"
$servers = Get-XAServer -EA 0 | sort-object FolderPath, ServerName

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Servers:"
	ForEach($server in $servers)
	{
		WriteWordLine 2 0 $server.ServerName
		WriteWordLine 0 1 "Product`t`t`t`t: " $server.CitrixProductName
		WriteWordLine 0 1 "Edition`t`t`t`t: " $server.CitrixEdition
		WriteWordLine 0 1 "Version`t`t`t`t: " $server.CitrixVersion
		WriteWordLine 0 1 "Service Pack`t`t`t: " $server.CitrixServicePack
		WriteWordLine 0 1 "Operating System Type`t`t: " -NoNewLine
		If($server.Is64Bit)
		{
			WriteWordLine 0 0 "64 bit"
		} 
		Else 
		{
			WriteWordLine 0 0 "32 bit"
		}
		WriteWordLine 0 1 "IP Address`t`t`t: " $server.IPAddresses
		WriteWordLine 0 1 "Logons`t`t`t`t: " -NoNewLine
		If($server.LogOnsEnabled)
		{
			WriteWordLine 0 0 "Enabled"
		} 
		Else 
		{
			WriteWordLine 0 0 "Disabled"
		}
		WriteWordLine 0 1 "Product Installation Date`t: " $server.CitrixInstallDate
		WriteWordLine 0 1 "Operating System Version`t: " $server.OSVersion -NoNewLine
		
		#is the server running server 2008?
		If($server.OSVersion.ToString().SubString(0,1) -eq "6")
		{
			$global:Server2008 = $True
		}

		WriteWordLine 0 0 " " $server.OSServicePack
		WriteWordLine 0 1 "Zone`t`t`t`t: " $server.ZoneName
		WriteWordLine 0 1 "Election Preference`t`t: " -nonewline
		switch ($server.ElectionPreference)
		{
			"Unknown"           {WriteWordLine 0 0 "Unknown"}
			"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"}
			"Preferred"         {WriteWordLine 0 0 "Preferred"}
			"DefaultPreference" {WriteWordLine 0 0 "Default Preference"}
			"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"}
			"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"}
			Default {WriteWordLine 0 0 "Server election preference could not be determined: $($server.ElectionPreference)"}
		}
		WriteWordLine 0 1 "Folder`t`t`t`t: " $server.FolderPath
		WriteWordLine 0 1 "Product Installation Path`t: " $server.CitrixInstallPath
		If($server.ICAPortNumber -gt 0)
		{
			WriteWordLine 0 1 "ICA Port Number`t`t: " $server.ICAPortNumber
		}
		$ServerConfig = Get-XAServerConfiguration -ServerName $Server.ServerName -EA 0
		If( $? )
		{
			WriteWordLine 0 1 "Server Configuration Data:"
			
			If($FarmOS -eq "2003")
			{
				$Text = "HDX Broadcast"
			}
			Else
			{
				$Text = "ICA"
			}
			
			If($ServerConfig.AcrUseFarmSettings)
			{
				WriteWordLine 0 2 "$($Text)\Auto Client Reconnect: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)\Auto Client Reconnect: Server is not using farm settings"
				If($ServerConfig.AcrEnabled)
				{
					WriteWordLine 0 3 "Reconnect automatically"
					WriteWordLine 0 4 "Log automatic reconnection attempts: " -nonewline
					If($ServerConfig.AcrLogReconnections)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				Else
				{
					WriteWordLine 0 3 "Require user authentication"
				}
			}
			WriteWordLine 0 2 "$($Text)\Browser\Create browser listener on UDP network: " -nonewline
			If($ServerConfig.BrowserUdpListener)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 2 "$($Text)\Browser\Server responds to client broadcast messages: " -nonewline
			If($ServerConfig.BrowserRespondToClientBroadcasts)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			If($ServerConfig.DisplayUseFarmSettings)
			{
				WriteWordLine 0 2 "$($Text)\Display: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)\Display: Server is not using farm settings"
				WriteWordLine 0 3 "Discard queued image that is replaced by another image: " -nonewline
				If($ServerConfig.DisplayDiscardQueuedImages)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 3 "Cache image to make scrolling smoother: " -nonewline
				If($ServerConfig.DisplayCacheImageForSmoothScrolling)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 3 "Maximum memory to use for each session's graphics (KB): " $ServerConfig.DisplayMaximumGraphicsMemory
				WriteWordLine 0 3 "Degradation bias: " 
				If($ServerConfig.DisplayDegradationBias -eq "Resolution")
				{
					WriteWordLine 0 4 "Degrade resolution first"
				}
				Else
				{
					WriteWordLine 0 4 "Degrade color depth first"
				}
				WriteWordLine 0 3 "Notify user of session degradation: " -nonewline
				If($ServerConfig.DisplayNotifyUser)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
			}
			If($ServerConfig.KeepAliveUseFarmSettings)
			{
				WriteWordLine 0 2 "$($Text)\Keep-Alive: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)\Keep-Alive: Server is not using farm settings"
				If($FarmOS -eq "2003")
				{
					WriteWordLine 0 3 "HDX Broadcast Keep-Alive time-out value seconds: " -NoNewLine
				}
				Else
				{
					WriteWordLine 0 3 "ICA Keep-Alive time-out value seconds: " -NoNewLine
				}
				If($ServerConfig.KeepAliveEnabled)
				{
					WriteWordLine 0 0 $ServerConfig.KeepAliveTimeout
				}
				Else
				{
					WriteWordLine 0 0 "Disabled"
				}
			}
			If($ServerConfig.PrinterBandwidth -eq -1)
			{
				If($FarmOS -eq "2003")
				{
					WriteWordLine 0 2 "$($Text)\Printer Bandwidth\Unlimited bandwidth"
				}
				Else
				{
					WriteWordLine 0 2 "$($Text)\Printer Bandwidth\Unlimited client printer bandwidth"
				}
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)\Printer Bandwidth\Limit bandwidth to use (kbps): " $ServerConfig.PrinterBandwidth
			}
			If($FarmOS -eq "2003")
			{
				If($ServerConfig.RemoteConsoleUseFarmSettings)
				{
					WriteWordLine 0 2 "$($Text)\Remote Console Connections: Server is using farm settings"
				}
				Else
				{
					WriteWordLine 0 2 "$($Text)\Remote Console Connections: Server is not using farm settings"
					WriteWordLine 0 3 "Remote connections to the console: " -nonewline
					If($ServerConfig.RemoteConsoleEnabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
			}
			
			#For server 2003, Isolation Environment data is not available

			If($ServerConfig.LicenseServerUseFarmSettings)
			{
				WriteWordLine 0 2 "License Server: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "License Server: Server is not using farm settings"
				WriteWordLine 0 3 "License server name: " $ServerConfig.LicenseServerName
				WriteWordLine 0 3 "License server port: " $ServerConfig.LicenseServerPortNumber
			}
			If($ServerConfig.HmrUseFarmSettings)
			{
				WriteWordLine 0 2 "Health Monitoring & Recovery: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "Health Monitoring & Recovery: Server is not using farm settings"
				WriteWordLine 0 3 "Run health monitoring tests on this server: " -nonewline
				If($ServerConfig.HmrEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If($ServerConfig.HmrEnabled)
				{
					$HMRTests = Get-XAHmrTest -ServerName $Server.ServerName -EA 0
					If( $? )
					{
						WriteWordLine 0 3 "Health Monitoring Tests:"
						ForEach($HMRTest in $HMRTests)
						{
							WriteWordLine 0 4 "Test Name`t: " $Hmrtest.TestName
							WriteWordLine 0 4 "Interval`t`t: " $Hmrtest.Interval
							WriteWordLine 0 4 "Threshold`t: " $Hmrtest.Threshold
							WriteWordLine 0 4 "Time-out`t: " $Hmrtest.Timeout
							WriteWordLine 0 4 "Test File Name`t: " $Hmrtest.FilePath
							If(![String]::IsNullOrEmpty($Hmrtest.Arguments))
							{
								WriteWordLine 0 4 "Arguments`t: " $Hmrtest.Arguments
							}
							WriteWordLine 0 4 "Recovery Action : " -nonewline
							switch ($Hmrtest.RecoveryAction)
							{
								"AlertOnly"                     {WriteWordLine 0 0 "Alert Only"}
								"RemoveServerFromLoadBalancing" {WriteWordLine 0 0 "Remove Server from load balancing"}
								"RestartIma"                    {WriteWordLine 0 0 "Restart IMA"}
								"ShutdownIma"                   {WriteWordLine 0 0 "Shutdown IMA"}
								"RebootServer"                  {WriteWordLine 0 0 "Reboot Server"}
								Default {WriteWordLine 0 0 "Recovery Action could not be determined: $($Hmrtest.RecoveryAction)"}
							}
							WriteWordLine 0 0 ""
						}
					}
					Else
					{
						WriteWordLine 0 0 "Health Monitoring & Reporting data could not be retrieved for server " $Server.ServerName
					}
				}
			}
			If($ServerConfig.CpuUseFarmSettings)
			{
				WriteWordLine 0 2 "CPU Utilization Management: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "CPU Utilization Management: Server is not using farm settings"
				WriteWordLine 0 3 "CPU Utilization Management: " -nonewline
				switch ($ServerConfig.CpuManagementLevel)
				{
					"NoManagement"  {WriteWordLine 0 0 "No CPU utilization management"}
					"Fair"          {WriteWordLine 0 0 "Fair sharing of CPU between sessions"}
					"ResourceBased" {WriteWordLine 0 0 "CPU Sharing based on Resource Allotments"}
					Default {WriteWordLine 0 0 "CPU Utilization Management could not be determined: $($Farm.CpuManagementLevel)"}
				}
			}
			If($ServerConfig.MemoryUseFarmSettings)
			{
				WriteWordLine 0 2 "Memory Optimization: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "Memory Optimization: Server is not using farm settings"
				WriteWordLine 0 3 "Memory Optimization: " -nonewline
				If($ServerConfig.MemoryOptimizationEnabled)
				{
					WriteWordLine 0 0 "Enabled"
				}
				Else
				{
					WriteWordLine 0 0 "Not Enabled"
				}
			}
			
			If($FarmOS -eq "2003")
			{
				$Text = "HDX Plug and Play"
			}
			Else
			{
				$Text = "XenApp"
			}
			
			If($ServerConfig.ContentRedirectionUseFarmSettings )
			{
				WriteWordLine 0 2 "$($Text)/Content Redirection: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)/Content Redirection: Server is not using farm settings"
				WriteWordLine 0 3 "Content redirection from server to client: " -nonewline
				If($ServerConfig.ContentRedirectionEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
			}
			#ShadowLoggingEnabled is not stored by Citrix
			#WriteWordLine 0 3 "HDX Plug and Play/Shadow Logging/Log shadowing sessions: " $ServerConfig.ShadowLoggingEnabled
			
			If($FarmOS -eq "2008")
			{
				If($ServerConfig.RemoteConsoleUseFarmSettings)
				{
					WriteWordLine 0 2 "$($Text)\Remote Console Connections: Server is using farm settings"
				}
				Else
				{
					WriteWordLine 0 2 "$($Text)\Remote Console Connections: Server is not using farm settings"
					WriteWordLine 0 3 "Remote connections to the console: " -nonewline
					If($ServerConfig.RemoteConsoleEnabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
			}

			If($ServerConfig.SnmpUseFarmSettings )
			{
				WriteWordLine 0 2 "SNMP: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "SNMP: Server is not using farm settings"
				# SnmpEnabled is not working
				WriteWordLine 0 3 "Send session traps to selected SNMP agent on all farm servers"
				WriteWordLine 0 4 "SNMP agent session traps:"
				WriteWordLine 0 5 "Logon`t`t`t: " -nonewline
				If($ServerConfig.SnmpLogOnEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 5 "Logoff`t`t`t: " -nonewline
				If($ServerConfig.SnmpLogOffEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 5 "Disconnect`t`t: " -nonewline
				If($ServerConfig.SnmpDisconnectEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 5 "Session limit per server`t: " -nonewline
				If($ServerConfig.SnmpLimitEnabled)
				{
					WriteWordLine 0 0 " " $ServerConfig.SnmpLimitPerServer
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
			}
			
			If($FarmOS -eq "2003")
			{
				$Text = "HDX 3D"
			}
			Else
			{
				$Text = "SpeedScreen"
			}
			
			If($ServerConfig.BrowserAccelerationUseFarmSettings )
			{
				WriteWordLine 0 2 "$($Text)/Browser Acceleration: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)/Browser Acceleration: Server is not using farm settings"
				WriteWordLine 0 3 "$($Text)/$($Text)Browser Acceleration: " -nonewline
				If($ServerConfig.BrowserAccelerationEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If($ServerConfig.BrowserAccelerationEnabled)
				{
					WriteWordLine 0 4 "Compress JPEG images to improve bandwidth: " -nonewline
					If($ServerConfig.BrowserAccelerationCompressionEnabled)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($ServerConfig.BrowserAccelerationCompressionEnabled)
					{
						WriteWordLine 0 4 "Image compression level: " $ServerConfig.BrowserAccelerationCompressionLevel
						WriteWordLine 0 4 "Adjust compression level based on available bandwidth: " -nonewline
						If($ServerConfig.BrowserAccelerationVariableImageCompression)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}
				}
			}
			
			If($FarmOS -eq "2003")
			{
				$Text = "HDX MediaStream"
			}
			Else
			{
				$Text = "SpeedScreen"
			}
			
			If($ServerConfig.FlashAccelerationUseFarmSettings )
			{
				WriteWordLine 0 2 "$($Text)/Flash: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)/Flash: Server is not using farm settings"
				If($FarmOS -eq "2003")
				{
					WriteWordLine 0 3 "Enable Flash for XenApp sessions: " -nonewline
				}
				Else
				{
					WriteWordLine 0 3 "Enable Adobe Flash Player: " -nonewline
				}
				If($ServerConfig.FlashAccelerationEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If($ServerConfig.FlashAccelerationEnabled)
				{
					switch ($ServerConfig.FlashAccelerationOption)
					{
						"RestrictedBandwidth" {WriteWordLine 0 3 "Restricted bandwidth connections"}
						"NoOptimization"      {WriteWordLine 0 3 "Do not optimize"}
						"AllConnections"      {WriteWordLine 0 3 "All connections"}
						Default {WriteWordLine 0 0 "Server-side acceleration could not be determined: $($ServerConfig.FlashAccelerationOption)"}
					}
				}
			}
			If($ServerConfig.MultimediaAccelerationUseFarmSettings )
			{
				WriteWordLine 0 2 "$($Text)/Multimedia Acceleration: Server is using farm settings"
			}
			Else
			{
				WriteWordLine 0 2 "$($Text)/Multimedia Acceleration: Server is not using farm settings"
				WriteWordLine 0 3 "Multimedia acceleration: " -nonewline
				If($ServerConfig.MultimediaAccelerationEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If($ServerConfig.MultimediaAccelerationEnabled)
				{
					If($ServerConfig.MultimediaAccelerationDefaultBuffer)
					{
						WriteWordLine 0 3 "Use the default buffer of 5 seconds"
					}
					Else
					{
						WriteWordLine 0 3 "Custom buffer time in seconds: " $ServerConfig.MultimediaAccelerationCustomBuffer
					}
				}
			}
			WriteWordLine 0 2 "Virtual IP/Enable virtual IP for this server: " -nonewline
			If($ServerConfig.VirtualIPEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 2 "Virtual IP/Use farm setting for IP address logging: " -nonewline
			If($ServerConfig.VirtualIPUseFarmLoggingSettings)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 2 "Virtual IP/Enable logging of IP address assignment and release: " -nonewline
			If($ServerConfig.VirtualIPLoggingEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 2 "Virtual IP/Enable virtual loopback for this server: " -nonewline
			If($ServerConfig.VirtualIPLoopbackEnabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			WriteWordLine 0 2 "XML Service/Trust requests sent to the XML service: " -nonewline
			If($ServerConfig.XmlServiceTrustRequests)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			If($Server2008)
			{
				If($ServerConfig.RestartsEnabled)
				{
					WriteWordLine 0 2 "Automatic restarts are enabled"
					WriteWordLine 0 3 "Restart server from: " $ServerConfig.RestartFrom
					WriteWordLine 0 3 "Restart frequency in days: " $ServerConfig.RestartFrequency
				}
				Else
				{
					WriteWordLine 0 2 "Automatic restarts are not enabled"
				}
			}
			WriteWordLine 0 0 ""
		}
		Else
		{
			WriteWordLine 0 0 "Server configuration data could not be retrieved for server " $Server.ServerName
		}
		#applications published to server
		$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			WriteWordLine 0 1 "Published applications:"
			ForEach($app in $Applications)
			{
				WriteWordLine 0 2 "Display name`t: " $app.DisplayName
				WriteWordLine 0 2 "Folder path`t: " $app.FolderPath
				WriteWordLine 0 0 ""
			}
		}

		#list citrix services
		Write-Verbose "`t`tTesting to see if $($server.ServerName) is online and reachable"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			Write-Verbose "`t`t$($server.ServerName) is online.  Citrix Services and Hotfix areas processed."
			Write-Verbose "`t`tProcessing Citrix services for server $($server.ServerName)"
			$services = get-service -ComputerName $server.ServerName -EA 0 | where-object {$_.DisplayName -like "*Citrix*"} | sort-object DisplayName
			WriteWordLine 0 1 "Citrix Services"
			Write-Verbose "`t`tCreate Word Table for Citrix services"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = $services.count + 1
			Write-Verbose "`t`tadd Citrix Services table to doc"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			$xRow = 1
			Write-Verbose "`t`tformat first row with column headings"
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Display Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Status"
			ForEach($Service in $Services)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $Service.DisplayName
				If($Service.Status -eq "Stopped")
				{
					$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
					$Table.Cell($xRow,2).Range.Font.Bold  = $True
					$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
				}
				$Table.Cell($xRow,2).Range.Text = $Service.Status
			}

			Write-Verbose "`t`tMove table of Citrix services to the right"
			$Table.Rows.SetLeftIndent(43,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			Write-Verbose "`t`treturn focus back to document"
			$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

			#move to the end of the current document
			Write-Verbose "`t`tmove to the end of the current document"
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			
			#Citrix hotfixes installed
			Write-Verbose "`t`tGet list of Citrix hotfixes installed"
			$hotfixes = Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | sort-object HotfixName
			If( $? -and $hotfixes )
			{
				$Rows = 1
				$Single_Row = (Get-Member -Type Property -Name Length -InputObject $hotfixes -EA 0) -eq $null
				If(-not $Single_Row)
				{
					$Rows = $Hotfixes.length
				}
				$Rows++
				
				Write-Verbose "`t`tnumber of hotfixes is $($Rows-1)"
				$HotfixArray = ""
				$HRP1Installed = $False
				WriteWordLine 0 0 ""
				WriteWordLine 0 1 "Citrix Installed Hotfixes:"
				Write-Verbose "`t`tCreate Word Table for Citrix Hotfixes"
				$TableRange = $doc.Application.Selection.Range
				$Columns = 5
				Write-Verbose "`t`tadd Citrix installed hotfix table to doc"
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				$xRow = 1
				Write-Verbose "`t`tformat first row with column headings"
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Font.Size = "10"
				$Table.Cell($xRow,1).Range.Text = "Hotfix"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Font.Size = "10"
				$Table.Cell($xRow,2).Range.Text = "Installed By"
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,3).Range.Font.Bold = $True
				$Table.Cell($xRow,3).Range.Font.Size = "10"
				$Table.Cell($xRow,3).Range.Text = "Install Date"
				$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,4).Range.Font.Bold = $True
				$Table.Cell($xRow,4).Range.Font.Size = "10"
				$Table.Cell($xRow,4).Range.Text = "Type"
				$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,5).Range.Font.Bold = $True
				$Table.Cell($xRow,5).Range.Font.Size = "10"
				$Table.Cell($xRow,5).Range.Text = "Valid"
				ForEach($hotfix in $hotfixes)
				{
					$xRow++
					$HotfixArray += $hotfix.HotfixName
					$InstallDate = $hotfix.InstalledOn.ToString()
					
					$Table.Cell($xRow,1).Range.Font.Size = "10"
					$Table.Cell($xRow,1).Range.Text = $hotfix.HotfixName
					$Table.Cell($xRow,2).Range.Font.Size = "10"
					$Table.Cell($xRow,2).Range.Text = $hotfix.InstalledBy
					$Table.Cell($xRow,3).Range.Font.Size = "10"
					$Table.Cell($xRow,3).Range.Text = $InstallDate.SubString(0,$InstallDate.IndexOf(" "))
					$Table.Cell($xRow,4).Range.Font.Size = "10"
					$Table.Cell($xRow,4).Range.Text = $hotfix.HotfixType
					$Table.Cell($xRow,5).Range.Font.Size = "10"
					$Table.Cell($xRow,5).Range.Text = $hotfix.Valid
				}
				Write-Verbose "`t`tMove table of Citrix installed hotfixes to the right"
				$Table.Rows.SetLeftIndent(43,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				Write-Verbose "`t`treturn focus back to document"
				$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

				#move to the end of the current document
				Write-Verbose "`t`tmove to the end of the current document"
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				WriteWordLine 0 0 ""
			}
		}
		Else
		{
			Write-Verbose "`t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped."
			WriteWordLine 0 0 "Server $($server.ServerName) was offline or unreachable at "(get-date).ToString()
			WriteWordLine 0 0 "The Citrix Services and Hotfix areas were skipped."
		}

		WriteWordLine 0 0 "" 
	}
}
Else 
{
	Write-warning "Server information could not be retrieved"
}
$servers = $null

write-verbose "Processing Zones"
$Zones = Get-XAZone -EA 0 | sort-object ZoneName
If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Zones:"
	ForEach($Zone in $Zones)
	{
		WriteWordLine 2 0 $Zone.ZoneName
		WriteWordLine 0 1 "Current Data Collector: " $Zone.DataCollector
		$Servers = Get-XAServer -ZoneName $Zone.ZoneName -EA 0 | sort-object ElectionPreference, ServerName
		If( $? )
		{		
			WriteWordLine 0 1 "Servers in Zone"
	
			ForEach($Server in $Servers)
			{
				WriteWordLine 0 2 "Server Name and Preference: " $server.ServerName -NoNewLine
				WriteWordLine 0 0  " - " -nonewline
				switch ($server.ElectionPreference)
				{
					"Unknown"           {WriteWordLine 0 0 "Unknown"}
					"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"}
					"Preferred"         {WriteWordLine 0 0 "Preferred"}
					"DefaultPreference" {WriteWordLine 0 0 "Default Preference"}
					"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"}
					"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"}
					Default {WriteWordLine 0 0 "Zone preference could not be determined: $($server.ElectionPreference)"}
				}
			}
		}
		Else
		{
			WriteWordLine 0 1 "Unable to enumerate servers in the zone"
		}
	}
}
Else 
{
	Write-warning "Zone information could not be retrieved"
}
$Servers = $null
$Zone = $null

#Process the nodes in the Advanced Configuration COnsole

write-verbose "Processing Load Evaluators"
#load evaluators
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | sort-object LoadEvaluatorName

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Load Evaluators:"
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		WriteWordLine 2 0 $LoadEvaluator.LoadEvaluatorName
		WriteWordLine 0 1 "Description: " $LoadEvaluator.Description
		
		If($LoadEvaluator.IsBuiltIn)
		{
			WriteWordLine 0 1 "Built-in Load Evaluator"
		} 
		Else 
		{
			WriteWordLine 0 1 "User created load evaluator"
		}
	
		If($LoadEvaluator.ApplicationUserLoadEnabled)
		{
			WriteWordLine 0 1 "Application User Load Settings"
			WriteWordLine 0 2 "Report full load when the # of users for this application =: " $LoadEvaluator.ApplicationUserLoad
			WriteWordLine 0 2 "Application: " $LoadEvaluator.ApplicationBrowserName
		}
	
		If($LoadEvaluator.ContextSwitchesEnabled)
		{
			WriteWordLine 0 1 "Context Switches Settings"
			WriteWordLine 0 2 "Report full load when the # of context switches per second is >= than: " $LoadEvaluator.ContextSwitches[1]
			WriteWordLine 0 2 "Report no load when the # of context switches per second is <= to: " $LoadEvaluator.ContextSwitches[0]
		}
	
		If($LoadEvaluator.CpuUtilizationEnabled)
		{
			WriteWordLine 0 1 "CPU Utilization Settings"
			WriteWordLine 0 2 "Report full load when the processor utilization % is > than: " $LoadEvaluator.CpuUtilization[1]
			WriteWordLine 0 2 "Report no load when the processor utilization % is <= to: " $LoadEvaluator.CpuUtilization[0]
		}
	
		If($LoadEvaluator.DiskDataIOEnabled)
		{
			WriteWordLine 0 1 "Disk Data I/O Settings"
			WriteWordLine 0 2 "Report full load when the total disk I/O in kbps > than: " $LoadEvaluator.DiskDataIO[1]
			WriteWordLine 0 2 "Report no load when the total disk I/O in kbps <= to: " $LoadEvaluator.DiskDataIO[0]
		}
	
		If($LoadEvaluator.DiskOperationsEnabled)
		{
			WriteWordLine 0 1 "Disk Operations Settings"
			WriteWordLine 0 2 "Report full load when the total # of R/W operations per second is > than: " $LoadEvaluator.DiskOperations[1]
			WriteWordLine 0 2 "Report no load when the total # of R/W operations per second is <= to: " $LoadEvaluator.DiskOperations[0]
		}
	
		If($LoadEvaluator.IPRangesEnabled)
		{
			WriteWordLine 0 1 "IP Range Settings"
			If($LoadEvaluator.IPRangesAllowed)
			{
				WriteWordLine 0 2 "Allow " -NoNewLine
			} 
			Else 
			{
				WriteWordLine 0 2 "Deny " -NoNewLine
			}
			WriteWordLine 0 2 "client connections from the listed IP Ranges"
			ForEach($IPRange in $LoadEvaluator.IPRanges)
			{
				WriteWordLine 0 4 "IP Address Ranges: " $IPRange
			}
		}

		If($LoadEvaluator.LoadThrottlingEnabled)
		{
			WriteWordLine 0 1 "Load Throttling Settings"
			WriteWordLine 0 2 "Impact of logons on load: " -nonewline
			switch ($LoadEvaluator.LoadThrottling)
			{
				"Unknown"    {WriteWordLine 0 0 "Unknown"}
				"Extreme"    {WriteWordLine 0 0 "Extreme"}
				"High"       {WriteWordLine 0 0 "High (Default)"}
				"MediumHigh" {WriteWordLine 0 0 "Medium High"}
				"Medium"     {WriteWordLine 0 0 "Medium"}
				"MediumLow"  {WriteWordLine 0 0 "Medium Low"}
				Default {WriteWordLine 0 0 "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"}
			}
		}
		
		If($LoadEvaluator.MemoryUsageEnabled)
		{
			WriteWordLine 0 1 "Memory Usage Settings"
			WriteWordLine 0 2 "Report full load when the memory usage is > than: " $LoadEvaluator.MemoryUsage[1]
			WriteWordLine 0 2 "Report no load when the memory usage is <= to: " $LoadEvaluator.MemoryUsage[0]
		}
	
		If($LoadEvaluator.PageFaultsEnabled)
		{
			WriteWordLine 0 1 "Page Faults Settings"
			WriteWordLine 0 2 "Report full load when the # of page faults per second is > than: " $LoadEvaluator.PageFaults[1]
			WriteWordLine 0 2 "Report no load when the # of page faults per second is <= to: " $LoadEvaluator.PageFaults[0]
		}
	
		If($LoadEvaluator.PageSwapsEnabled)
		{
			WriteWordLine 0 1 "Page Swaps Settings"
			WriteWordLine 0 2 "Report full load when the # of page swaps per second is > than: " $LoadEvaluator.PageSwaps[1]
			WriteWordLine 0 2 "Report no load when the # of page swaps per second is <= to: " $LoadEvaluator.PageSwaps[0]
		}
	
		If($LoadEvaluator.ScheduleEnabled)
		{
			WriteWordLine 0 1 "Scheduling Settings"
			WriteWordLine 0 2 "Sunday Schedule`t: " $LoadEvaluator.SundaySchedule
			WriteWordLine 0 2 "Monday Schedule`t: " $LoadEvaluator.MondaySchedule
			WriteWordLine 0 2 "Tuesday Schedule`t: " $LoadEvaluator.TuesdaySchedule
			WriteWordLine 0 2 "Wednesday Schedule`t: " $LoadEvaluator.WednesdaySchedule
			WriteWordLine 0 2 "Thursday Schedule`t: " $LoadEvaluator.ThursdaySchedule
			WriteWordLine 0 2 "Friday Schedule`t`t: " $LoadEvaluator.FridaySchedule
			WriteWordLine 0 2 "Saturday Schedule`t: " $LoadEvaluator.SaturdaySchedule
		}

		If($LoadEvaluator.ServerUserLoadEnabled)
		{
			WriteWordLine 0 1 "Server User Load Settings"
			WriteWordLine 0 2 "Report full load when the # of server users equals: " $LoadEvaluator.ServerUserLoad
		}
	
		WriteWordLine 0 0 ""
	}
}
Else 
{
	Write-warning "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $null

write-verbose "Processing Policies"
$Policies = Get-XAPolicy -EA 0 | sort-object PolicyName
If( $? -and $Policies)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Policies:"
	ForEach($Policy in $Policies)
	{
		Write-Verbose "`t`tProcessing policy $($Policy.PolicyName)"
		WriteWordLine 2 0 $Policy.PolicyName
		WriteWordLine 0 1 "Description: " $Policy.Description
		WriteWordLine 0 1 "Enabled: " $Policy.Enabled
		WriteWordLine 0 1 "Priority: " $Policy.Priority

		$filter = Get-XAPolicyFilter -PolicyName $Policy.PolicyName -EA 0

		If( $? )
		{
			If($Filter)
			{
				WriteWordLine 0 1 "Policy Filters:"
				
				If($Filter.AccessControlEnabled)
				{
					If($Filter.AllowConnectionsThroughAccessGateway)
					{
						WriteWordLine 0 2 "Apply to connections made through Access Gateway"
						If($Filter.AccessSessionConditions)
						{
							WriteWordLine 0 3 "Any connection that meets any of the following filters"
							ForEach($Condition in $Filter.AccessSessionConditions)
							{
								$Colon = $Condition.IndexOf(":")
								$CondName = $Condition.SubString(0,$Colon)
								$CondFilter = $Condition.SubString($Colon+1)
								WriteWordLine 0 4 "Access Gateway Farm Name: " $CondName -NoNewLine
								WriteWordLine 0 0 " Filter Name: " $CondFilter
							}
						}
						Else
						{
							WriteWordLine 0 3 "Any connection"
						}
						WriteWordLine 0 3 "Apply to all other connections: " -nonewline
						If($Filter.AllowOtherConnections)
						{
							WriteWordLine 0 0 "Yes"
						}
						Else
						{
							WriteWordLine 0 0 "No"
						}
					}
					Else
					{
						WriteWordLine 0 2 "Do not apply to connections made through Access Gateway"
					}
				}
				If($Filter.ClientIPAddressEnabled)
				{
					WriteWordLine 0 2 "Apply to all client IP addresses: " -nonewline
					If($Filter.ApplyToAllClientIPAddresses)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($Filter.AllowedIPAddresses)
					{
						WriteWordLine 0 2 "Allowed IP Addresses:"
						ForEach($Allowed in $Filter.AllowedIPAddresses)
						{
							WriteWordLine 0 3 $Allowed
						}
					}
					If($Filter.DeniedIPAddresses)
					{
						WriteWordLine 0 2 "Denied IP Addresses:"
						ForEach($Denied in $Filter.DeniedIPAddresses)
						{
							WriteWordLine 0 3 $Denied
						}
					}
				}
				If($Filter.ClientNameEnabled)
				{
					WriteWordLine 0 2 "Apply to all client names: " -nonewline
					If($Filter.ApplyToAllClientNames)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($Filter.AllowedClientNames)
					{
						WriteWordLine 0 2 "Allowed Client Names:"
						ForEach($Allowed in $Filter.AllowedClientNames)
						{
							WriteWordLine 0 3 $Allowed
						}
					}
					If($Filter.DeniedClientNames)
					{
						WriteWordLine 0 2 "Denied Client Names:"
						ForEach($Denied in $Filter.DeniedClientNames)
						{
							WriteWordLine 0 3 $Denied
						}
					}
				}
				If($Filter.ServerEnabled)
				{
					If($Filter.AllowedServerNames)
					{
						WriteWordLine 0 2 "Allowed Server Names:"
						ForEach($Allowed in $Filter.AllowedServerNames)
						{
							WriteWordLine 0 3 $Allowed
						}
					}
					If($Filter.DeniedServerNames)
					{
						WriteWordLine 0 2 "Denied Server Names:"
						ForEach($Denied in $Filter.DeniedServerNames)
						{
							WriteWordLine 0 3 $Denied
						}
					}
					If($Filter.AllowedServerFolders)
					{
						WriteWordLine 0 2 "Allowed Server Folders:"
						ForEach($Allowed in $Filter.AllowedServerFolders)
						{
							WriteWordLine 0 3 $Allowed
						}
					}
					If($Filter.DeniedServerFolders)
					{
						WriteWordLine 0 2 "Denied Server Folders:"
						ForEach($Denied in $Filter.DeniedServerFolders)
						{
							WriteWordLine 0 3 $Denied
						}
					}
				}
				If($Filter.AccountEnabled)
				{
					WriteWordLine 0 2 "Apply to all explicit (non-anonymous) users: " -nonewline
					If($Filter.ApplyToAllExplicitAccounts)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					WriteWordLine 0 2 "Apply to anonymous users: " -nonewline
					If($Filter.ApplyToAnonymousAccounts)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($Filter.AllowedAccounts)
					{
						WriteWordLine 0 2 "Allowed Accounts:"
						ForEach($Allowed in $Filter.AllowedAccounts)
						{
							WriteWordLine 0 3 $Allowed
						}
					}
					If($Filter.DeniedAccounts)
					{
						WriteWordLine 0 2 "Denied Accounts:"
						ForEach($Denied in $Filter.DeniedAccounts)
						{
							WriteWordLine 0 3 $Denied
						}
					}
				}
			}
			Else
			{
				WriteWordLine 0 1 "No filter information"
			}
		}
		Else
		{
			WriteWordLine 0 1 "Unable to retrieve Filter settings"
		}

		$Global:Settings = Get-XAPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0

		If( $? )
		{
			WriteWordLine 0 1 "Policy Settings:"
			ForEach($Setting in $Settings)
			{
				If($FarmOS -eq "2003")
				{
					Process2003Policies
				}
				Else
				{
					Process2008Policies
				}
			}
		}
		Else
		{
			WriteWordLine 0 1 "Unable to retrieve settings"
		}
	
		$Global:Settings = $null
		$Filter = $null
	}
}
Else 
{
	Write-warning "Citrix Policy information could not be retrieved."
}
$Policies = $null

write-verbose "Processing Print Drivers"
#printer drivers
$PrinterDrivers = Get-XAPrinterDriver -EA 0 | sort-object DriverName

If( $? -and $PrinterDrivers)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Print Drivers:"
	ForEach($PrinterDriver in $PrinterDrivers)
	{
		WriteWordLine 0 1 "Driver`t: " $PrinterDriver.DriverName
		WriteWordLine 0 1 "Platform: " $PrinterDriver.OSVersion
		WriteWordLine 0 1 "64 bit?`t: " $PrinterDriver.Is64Bit
		WriteWordLine 0 0 ""
	}
}
Else 
{
	Write-warning "Printer driver information could not be retrieved"
}

$PrintDrivers = $null

write-verbose "Processing Printer Driver Mappings"
#printer driver mappings
$PrinterDriverMappings = Get-XAPrinterDriverMapping -EA 0 | sort-object ClientDriverName

If( $? -and $PrinterDriverMappings)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Print Driver Mappings:"
	ForEach($PrinterDriverMapping in $PrinterDriverMappings)
	{
		WriteWordLine 0 1 "Client Driver: " $PrinterDriverMapping.ClientDriverName
		WriteWordLine 0 1 "Server Driver: " $PrinterDriverMapping.ServerDriverName
		WriteWordLine 0 1 "Platform: " $PrintDriverMapping.OSVersion
		WriteWordLine 0 1 "64 bit? : " $PrinterDriverMapping.Is64Bit
		WriteWordLine 0 0 ""
	}
}
Else 
{
	Write-warning "Printer driver mapping information could not be retrieved"
}

$PrintDriverMappings = $null

If( $Global:ConfigLog)
{
	write-verbose "Processing the Configuration Logging Report"
	#Configuration Logging report
	#only process if $Global:ConfigLog = $True and .\XA5ConfigLog.udl file exists
	#build connection string for Microsoft SQL Server
	#User ID is account that has access permission for the configuration logging database
	#Initial Catalog is the name of the Configuration Logging SQL Database
	If ( Test-Path .\XA5ConfigLog.udl )
	{
		$ConnectionString = Get-Content .\xa5configlog.udl | select-object -last 1
		$ConfigLogReport = get-XAConfigurationLog -connectionstring $ConnectionString -EA 0

		If( $? -and $ConfigLogReport)
		{
			$selection.InsertNewPage()
			WriteWordLine 1 0 "Configuration Log Report:"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				WriteWordLine 0 1 "Date`t`t`t: " $ConfigLogItem.Date
				WriteWordLine 0 1 "Account`t`t: " $ConfigLogItem.Account
				WriteWordLine 0 1 "Change description`t: " $ConfigLogItem.Description
				WriteWordLine 0 1 "Type of change`t`t: " $ConfigLogItem.TaskType
				WriteWordLine 0 1 "Type of item`t`t: " $ConfigLogItem.ItemType
				WriteWordLine 0 1 "Name of item`t`t: " $ConfigLogItem.ItemName
				WriteWordLine 0 0 ""
			}
		} 
		Else 
		{
			WriteWordLine 0 0 "Configuration log report could not be retrieved"
		}
		$ConfigLogReport = $null
	}
	Else 
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Configuration Logging is enabled but the XA5ConfigLog.udl file was not found"
	}
}

write-verbose "Finishing up Word document"
#end of document processing
#Update document properties

If($CoverPagesExist)
{
	write-verbose "Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "XenApp 5 Farm Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp=$doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab=$cp.DocumentElement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract="Citrix XenApp 5 Inventory for $CompanyName"
	$ab.Text=$abstract

	$ab=$cp.DocumentElement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract=( Get-Date -Format d ).ToString()
	$ab.Text=$abstract

	write-verbose "Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
}

#the $saveFormat below passes StrictMode 2
#I found this at the following two links
#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
write-verbose "Save and Close document and Shutdown Word"
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
$doc.SaveAs([REF]$filename, [ref]$SaveFormat)
$doc.Close()
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word
[gc]::collect() 
[gc]::WaitForPendingFinalizers()

