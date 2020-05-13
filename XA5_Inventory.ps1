#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mcbogo on Twitter
#The original script was designed to be run on a XenApp 6 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from original script for XenApp 5
#originally released to the Citrix community on October 3, 2011
#updated October 9, 2011.  Added CPU Utilization Management, Memory Optimization and Health Monitoring & Recovery

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com

{
	Param( [int]$tabs = 0, [string]$name = ’’, [string]$value = ’’, [string]$newline = “`n”, [switch]$nonewline )

	While( $tabs –gt 0 ) { $global:output += “`t”; $tabs--; }

	If( $nonewline )
	{
		$global:output += $name + $value
	}
	Else
	{
		$global:output += $name + $value + $newline
	}
}

$farm = Get-XAFarm -EA 0
If( $? )
{
	#first check to make sure this is a XenApp 5 farm
	If($Farm.ServerVersion.ToString().SubString(0,1) -ne "6")
	{
		#this is a XenApp 5 farm, script can proceed
	}
	Else
	{
		#this is not a XenApp 5 farm, script cannot proceed
		write-warning "This script is designed for XenApp 5 and should not be run on XenApp 6.x"
		Return 1
	}
} 
$farm = $null
$global:output = ""

#process the nodes in the Delivery Services Console (XA5/2003) and the Access Management Console (XA5/2008)

# Get farm information
$global:Server2008 = $False
$global:ConfigLog = $False
$farm = Get-XAFarmConfiguration -EA 0

If( $? )
{
	line 0 "Farm: "$farm.FarmName
	
	line 1 "Farm-wide"

	line 2 "Connection Access Controls"
	line 3 "Connection access controls"
	line 4 $Farm.ConnectionAccessControls

	line 2 "Connection Limits" 
	line 3 "Connections per user"
	line 4 "Maximum connections per user: " -NoNewLine
	If($Farm.ConnectionLimitsMaximumPerUser -eq -1)
	{
		line 0 "No limit set"
	}
	Else
	{
		line 0 $Farm.ConnectionLimitsMaximumPerUser
	}
	If($Farm.ConnectionLimitsEnforceAdministrators)
	{
		line 5 "Enforce limit on administrators"
	}
	Else
	{
		line 5 "Do not enforce limit on administrators"
	}

	line 4 "Logging"

	If($Farm.ConnectionLimitsLogOverLimits)
	{
		line 5 "Log over-the-limit denials"
	}
	Else
	{
		line 5 "Do not log over-the-limit denials"
	}

	line 2 "Health Monitoring & Recovery"
	line 3 "Limit server for load balancing"
	line 4 "Limit servers (%): " $Farm.HmrMaximumServerPercent

	line 2 "Configuration Logging"
	If($Farm.ConfigLogEnabled)
	{
		$global:ConfigLog = $True

		line 3 "Database configuration"
		line 4 "Database type: " $Farm.ConfigLogDatabaseType
		If($Farm.ConfigLogDatabaseAuthenticationMode -eq "Native")
		{
			line 4 "Use SQL Server authentication"
		}
		Else
		{
			line 4 "Use Windows integrated security"
		}

		line 4 "Connection String: " -NoNewLine

		$StringMembers = "`n`t`t`t`t`t" + $Farm.ConfigLogDatabaseConnectionString.replace(";","`n`t`t`t`t`t")
		
		line 4 $StringMembers -NoNewLine
		line 0 "User name=" $Farm.ConfigLogDatabaseUserName

		line 3 "Log tasks"
		line 4 "Log administrative tasks to logging database: " $Farm.ConfigLogEnabled
		line 5 "Disconnected database"
		line 6 "Allow changes to the farm when database is disconnected: " $Farm.ConfigLogChangesWhileDisconnectedAllowed
		line 3 "Clearing log"
		line 4 "Require administrators to enter database credentials before clearing the log: " $Farm.ConfigLogCredentialsOnClearLogRequired
	}
	Else
	{
		line 3 "Configuration logging is not enabled"
	}
		
	line 2 "Memory/CPU"

	line 3 "Exclude Applications"
	line 4 "Applications that memory optimization ignores: "
	If($Farm.MemoryOptimizationExcludedApplications)
	{
		ForEach($App in $Farm.MemoryOptimizationExcludedApplications)
		{
			line 5 $App
		}
	}
	Else
	{
		line 5 "No applications are listed"
	}

	line 3 "Optimization Interval"
	line 4 "Optimization interval: " $Farm.MemoryOptimizationScheduleType

	If($Farm.MemoryOptimizationScheduleType -eq "Weekly")
	{
		line 5 "Day of week: " + $Farm.MemoryOptimizationScheduleDayOfWeek
	}
	If($Farm.MemoryOptimizationScheduleType -eq "Monthly")
	{
		line 5 "Day of month: " + $Farm.MemoryOptimizationScheduleDayOfMonth
	}

	line 5 "Optimization time: " $Farm.MemoryOptimizationScheduleTime
	line 4 "Memory optimization user"
	If($Farm.MemoryOptimizationLocalSystemAccountUsed)
	{
		line 4 "Use local system account"
	}
	Else
	{
		line 4 "Account: " + $Farm.MemoryOptimizationUser
	}
	
	line 2 "XenApp"
	line 3 "General"
	line 4 "Respond to client broadcast messages"
	line 5 "Data collectors: " $Farm.RespondDataCollectors
	line 5 "RAS servers: " $Farm.RespondRasServers
	line 4 "Client time zones"
	line 5 "Use client's local time: " $Farm.ClientLocalTimeEnabled
	line 6 "Estimate local time for clients: " $Farm.ClientLocalTimeEstimationEnabled
	line 4 "Citrix XML Service"
	line 5 "XML Service DNS address resolution: " $Farm.DNSAddressResolution
	line 4 "Novell Directory Services"
	line 5 "NDS preferred tree: " -NoNewLine
	If($Farm.NdsPreferredTree)
	{
		line 0 $Farm.NdsPreferredTree
	}
	Else
	{
		line 0 "No NDS Tree entered"
	}
	line 4 "Enable 32 bit icon color depth"
	line 5 "Enable: " $Farm.EnhancedIconEnabled

	line 3 "Shadow Policies"
	line 4 "Shadow policies"
	line 5 "Merge shadowers in multiple policies: " $Farm.ShadowPoliciesMerge

	line 2 "HDX Broadcast"
	line 3 "Session Reliability"
	line 4 "Session reliability"
	line 5 "Allow users to view sessions during broken connection: " $Farm.SessionReliabilityEnabled
	line 5 "Port number (default 2598): " $Farm.SessionReliabilityPort
	line 5 "Seconds to keep sessions active: " $Farm.SessionReliabilityTimeout

	line 2 "Citrix Streaming Server"
	line 3 "Log Citrix Streaming Server application events"
	line 4 "Log application events to event log: " $Farm.StreamingLogEvents
	line 3 "Trust Citrix Streaming Clients"
	line 4 "Trust Citrix Delivery Clients: " $Farm.StreamingTrustCLient

	line 2 "Virtual IP"
	line 3 "Address Configuration"
	line 4 "Virtual IP address ranges:"

	$VirtualIPs = Get-XAVirtualIPRange
	If($? -and $VirtualIPs)
	{
		ForEach($VirtualIP in $VirtualIPs)
		{
			line 5 "IP Range: " $VirtualIP
		}
	}
	Else
	{
		line 5 "No virtual IP address range defined"
	}
	$VirtualIPs = $Null

	line 4 "Enable logging of IP address assignment and release: " $Farm.VirtualIPLoggingEnabled
	line 3 "Process Configuration"
	line 4 "Virtual IP Processes"
	If($Farm.VirtualIPProcess)
	{
		line 5 "Monitor the following processes:"
		ForEach($Process in $Farm.VirtualIPProcess)
		{
			line 6 "Process: " $Process
		}
	}
	Else
	{
		Line 5 "No virtual IP processes defined"
	}
	line 4 "Virtual Loopback Processes"
	If($Farm.VirtualIPLoopbackProcesses)
	{
		line 5 "Monitor the following processes:"
		ForEach($Process in $Farm.VirtualIPLoopbackProcesses)
		{
			line 6 "Process: " $Process
		}
	}
	Else
	{
		Line 5 "No virtual IP Loopback processes defined"
	}
		
	line 1 "Server Default"
	line 2 "HDX Broadcast"
	line 3 "Auto Client Reconnect"
	line 4 "Auto client reconnect"
	If($Farm.AcrEnabled)
	{
		line 5 "Reconnect automatically"
		line 5 "Log automatic reconnection attempts: " -NoNewLine

		If($Farm.AcrLogReconnections)
		{
			line 0 "Enabled"
		}
		Else
		{
			line 0 "Disabled"
		}
	}
	Else
	{
		line 5 "Require user authentication"
	}
	
	line 3 "Display"
	line 4 "HDX Broadcast Display"
	line 5 "Discard queue image"
	line 6 "Discard queued image that is replaced by another image: " $Farm.DisplayDiscardQueuedImages
	line 5 "Cache image"
	line 6 "Cache image to make scrolling smoother: " $Farm.DisplayCacheImageForSmoothScrolling
	line 5 "Maximum memory to use for each session's graphics (KB): " $Farm.DisplayMaximumGraphicsMemory
	line 5 "Degradation bias"
	If($Farm.DisplayDegradationBias -eq "Resolution")
	{
		line 6 "Degrade resolution first"
	}
	Else
	{
		line 6 "Degrade color depth first"
	}
	line 5 "Notify user of session degradation: " $Farm.DisplayNotifyUser

	line 3 "Keep-Alive"
	line 4 "Keep-Alive"
	If($Farm.KeepAliveEnabled)
	{
		line 5 "HDX Broadcast Keep-Alive time-out value (1-3600 seconds): " $Farm.KeepAliveTimeout
	}
	Else
	{
		line 5 "HDX Broadcast Keep-Alive not enabled"
	}
	
	line 3 "Remote Console Connections"
	line 4 "Remote console cOnnections"
	line 5 "Remove connections to the console: " $Farm.RemoteConsoleEnabled
	
	line 2 "License Server"
	line 3 "License server"
	line 4 "Name: " $Farm.LicenseServerName
	line 4 "Port number (default 27000): " $Farm.LicenseServerPortNumber
	
	line 2 "Memory/CPU"
	line 3 "CPU Utilization Management: " -NoNewLine
	If($Farm.CpuManagementLevel.ToString() -eq "255")
	{
		line 0 "Cannot be determined for XenApp 5 on Windows Server 2003"
	}
	Else
	{
		line 0 "" $Farm.CpuManagementLevel
	}
	line 3 "Memory Optimization: " $Farm.MemoryOptimizationEnabled
	
	line 2 "Health Monitoring & Recovery"
	If($Farm.HmrEnabled)
	{
		$HmrTests = Get-XAHmrTest -EA 0 | Sort-Object TestName
		If($?)
		{
			ForEach($HmrTest in $HmrTests)
			{
				line 3 "Test Name: " $Hmrtest.TestName
				line 3 "Interval: " $Hmrtest.Interval
				line 3 "Threshold: " $Hmrtest.Threshold
				line 3 "Time-out: " $Hmrtest.Timeout
				line 3 "Test File Name: " $Hmrtest.FilePath
				If($Hmrtest.Arguments)
				{
					line 4 "Arguments: " $Hmrtest.Arguments
				}
				line 3 "Recovery Action: " $Hmrtest.RecoveryAction
				line 3 "Test Description: " $Hmrtest.Description
				line 0 ""
			}
		}
		Else
		{
			line 3 "Health Monitoring & Recovery Tests could not be retrieved"
		}
	}
	Else
	{
		line 3 "Health Monitoring & Recovery is not enabled"
	}

	line 2 "HDX Play and Play"
	line 3 "Content Redirection"
	line 4 "Content redirection from server to client"
	line 5 "Content redirection from server to client: " $Farm.ContentRedirectionEnabled

	line 2 "SNMP"
	line 3 "SNMP"
	If($Farm.SnmpEnabled)
	{
		line 4 "Send session traps to selected SNMP agent on all farm servers"
		line 5 "SNMP agent session traps"
		line 6 "Logon: " $Farm.SnmpLogonEnabled
		line 6 "Logoff: " $Farm.SnmpLogoffEnabled
		line 6 "Disconnect: " $Farm.SnmpDisconnectEnabled
		line 6 "Session limit per server: " $Farm.SnmpLimitEnabled -NoNewLine
		line 0 " " SnmpLimitPerServer
	}
	Else
	{
		line 4 "SNMP is not enabled"
	}

	line 2 "HDX 3D"
	line 3 "Browser Acceleration"
	line 4 "HDX 3D Browser Acceleration"
	If($Farm.BrowserAccelerationEnabled)
	{
		line 5 "HDX 3D Browser Acceleration is enabled"
		If($Farm.BrowserAccelerationCompressionEnabled)
		{
			line 6 "Compress JPEG images to improve bandwidth"
			line 7 "Image compression levels: " $Farm.BrowserAccelerationCompressionLevel
			If($Farm.BrowserAccelerationVariableImageCompression)
			{
				line 7 "Adjust compression level based on available bandwidth"
			}
			Else
			{
				line 7 "Do not adjust compression level based on available bandwidth"
			}
		}
		Else
		{
			line 6 "Do not compress JPEG images to improve bandwidth"
		}
	}
	Else
	{
		line 5 "HDX 3D Browser Acceleration is disabled"
	}
	
	line 2 "HDX Mediastream"
	line 3 "Flash"
	If($Farm.FlashAccelerationEnabled)
	{
		line 4 "Enable Flash for XenApp sessions"
		line 5 "Server-side acceleration: " $Farm.FlashAccelerationOption
	}
	Else
	{
		line 4 "Flash is not enabled for XenApp sessions"
	}
	line 3 "Multimedia Acceleration"
	If($Farm.MultimediaAccelerationEnabled)
	{
		Line 4 "Multimedia Acceleration is enabled"
		line 5 "Multimedia Acceleration (Network buffering)"
		If($Farm.MultimediaAccelerationDefaultBuffer)
		{
			line 6 "Use the default buffer of 5 seconds (recommended)"
		}
		Else
		{
			Line 6 "Custom buffer in seconds (1-10): " $Farm.MultimediaAccelerationCustomBuffer
		}
	}
	Else
	{
		Line 4 "Multimedia Acceleration is disabled"
	}
	
	line 1 "Offline Access"
	line 2 "Users"
	If($Farm.OfflineAccounts)
	{
		line 3 "Configured users:"
		ForEach($User in $Farm.OfflineAccounts)
		{
			line 4 $User
		}
	}
	Else
	{
		line 3 "No users configured"
	}

	line 2 "Offline License Settings"
	line 3 "License period (2-365 days): " $Farm.OfflineLicensePeriod

} 
Else 
{
	line 0 "Farm information could not be retrieved"
}
write-output $global:output
$farm = $null
$global:output = $null

$Administrators = Get-XAAdministrator -EA 0 | sort-object AdministratorName

If( $? )
{
	line 0 ""
	line 0 "Administrators:"
	ForEach($Administrator in $Administrators)
	{
		line 0 ""
		line 1 "Administrator name: "$Administrator.AdministratorName
		line 1 "Administrator type: "$Administrator.AdministratorType -nonewline
		line 0 " Administrator"
		line 1 "Administrator account is " -NoNewLine
		If($Administrator.Enabled)
		{
			line 0 "Enabled" 
		} 
		Else
		{
			line 0 "Disabled" 
		}
		line 1 "Alert Contact Details"
		line 2 "E-mail: " $Administrator.EmailAddress
		line 2 "SMS Number: " $Administrator.SmsNumber
		line 2 "SMS Gateway: " $Administrator.SmsGateway
		If ($Administrator.AdministratorType -eq "Custom") 
		{
			line 1 "Farm Privileges:"
			ForEach($farmprivilege in $Administrator.FarmPrivileges) 
			{
				line 2 $farmprivilege
			}
	
			line 1 "Folder Privileges:"
			ForEach($folderprivilege in $Administrator.FolderPrivileges) 
			{
				$test = $folderprivilege.ToString()
				$folderlabel = $test.substring(0, $test.IndexOf(":") + 1)
				line 2 $folderlabel
				$test1 = $test.substring($test.IndexOf(":") + 1)
				$folderpermissions = $test1.replace(",","`n`t`t`t")
				line 3 $folderpermissions
			}
		}		
	
	write-output $global:output
	$global:output = $null
	}
}
Else 
{
	line 0 "Administrator information could not be retrieved"
	write-output $global:output
}

$Administrators = $null
$global:outout = $null

$Applications = Get-XAApplication -EA 0 | sort-object FolderPath, DisplayName

If( $? -and $Applications)
{
	line 0 ""
	line 0 "Applications:"
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
		line 0 ""
		line 1 "Display name: " $Application.DisplayName
		line 2 "Application name (Browser name): " $Application.BrowserName
		line 2 "Disable application: " -NoNewLine
		If ($Application.Enabled) 
		{
		  line 0 "False"
		} 
		Else
		{
		  line 0 "True"
		}
		line 2 "Hide disabled application: " $Application.HideWhenDisabled
		line 2 "Application description: " $Application.Description
	
		#type properties
		line 2 "Application Type: " $Application.ApplicationType
		line 2 "Folder path: " $Application.FolderPath
		line 2 "Content Address: " $Application.ContentAddress
	
		#if a streamed app
		If($streamedapp)
		{
			line 2 "Citrix streaming application profile address: " $Application.ProfileLocation
			line 2 "Application to launch from the Citrix streaming application profile: " $Application.ProfileProgramName
			line 2 "Extra command line parameters: " $Application.ProfileProgramArguments
			#if streamed, Offline access properties
			If($Application.OfflineAccessAllowed)
			{
				line 2 "Enable offline access: " $Application.OfflineAccessAllowed
			}
			If($Application.CachingOption)
			{
				line 2 "Cache preference: " $Application.CachingOption
			}
		}
		
		If(!$streamedapp)
		{
			#location properties
			line 2 "Command line: " $Application.CommandLineExecutable
			line 2 "Working directory: " $Application.WorkingDirectory
			
			#servers properties
			If($AppServerInfoResults)
			{
				line 2 "Servers:"
				ForEach($servername in $AppServerInfo.ServerNames)
				{
					line 3 $servername
				}
			}
			Else
			{
				line 3 "Unable to retrieve a list of Servers for this application"
			}
		}
	
		#users properties
		If($Application.AnonymousConnectionsAllowed)
		{
			line 2 "Allow anonymous users: " $Application.AnonymousConnectionsAllowed
		}
		Else
		{
			If($AppServerInfoResults)
			{
				line 2 "Users:"
				ForEach($user in $AppServerInfo.Accounts)
				{
					line 3 $user
				}
			}
			Else
			{
				line 3 "Unable to retrieve a list of Users for this application"
			}
		}
	
		#shortcut presentation properties
		#application icon is ignored
		line 2 "Client application folder: " $Application.ClientFolder
		If($Application.AddToClientStartMenu)
		{
			line 2 "Add to client's start menu: " $Application.AddToClientStartMenu
		}
		If($Application.StartMenuFolder)
		{
			line 2 "Start menu folder: " $Application.StartMenuFolder
		}
		If($Application.AddToClientDesktop)
		{
			line 2 "Add shortcut to the client's desktop: " $Application.AddToClientDesktop
		}
	
		#access control properties
		If($Application.ConnectionsThroughAccessGatewayAllowed)
		{
			line 2 "Allow connections made through AGAE: " $Application.ConnectionsThroughAccessGatewayAllowed
		}
		If($Application.OtherConnectionsAllowed)
		{
			line 2 "Any connection: " $Application.OtherConnectionsAllowed
		}
		If($Application.AccessSessionConditionsEnabled)
		{
			line 2 "Any connection that meets any of the following filters: " $Application.AccessSessionConditionsEnabled
			line 2 "Access Gateway Filters:"
			ForEach($filter in $Application.AccessSessionConditions)
			{
				line 3 $filter
			}
		}
	
		#content redirection properties
		If($AppServerInfoResults)
		{
			If($AppServerInfo.FileTypes)
			{
				line 2 "File type associations:"
				ForEach($filetype in $AppServerInfo.FileTypes)
				{
					line 3 $filetype
				}
			}
			Else
			{
				line 2 "No File Type Associations exist for this application"
			}
		}
		Else
		{
			line 2 "Unable to retrieve the list of File Type Associations for this application"
		}
	
		#if streamed app, Alternate profiles
		If($streamedapp)
		{
			If($Application.AlternateProfiles)
			{
				line 2 "Primary application profile location: " $Application.AlternateProfiles
			}
		
			#if streamed app, User privileges properties
			If($Application.RunAsLeastPrivilegedUser)
			{
				line 2 "Run application as a least-privileged user account: " $Application.RunAsLeastPrivilegedUser
			}
		}
	
		#limits properties
		line 2 "Limit instances allowed to run in server farm: " -NoNewLine

		If($Application.InstanceLimit -eq -1)
		{
			line 0 "No limit set"
		}
		Else
		{
			line 0 $Application.InstanceLimit
		}
	
		line 2 "Allow only one instance of application for each user: " -NoNewLine
	
		If ($Application.MultipleInstancesPerUserAllowed) 
		{
			line 0 "False"
		} 
		Else
		{
			line 0 "True"
		}
	
		If($Application.CpuPriorityLevel)
		{
			line 2 "Application importance: " $Application.CpuPriorityLevel
		}
		
		#client options properties
		If($Application.AudioRequired)
		{
			line 2 "Enable legacy audio: " $Application.AudioRequired
		}
		If($Application.AudioType)
		{
			line 2 "Minimum requirement: " $Application.AudioType
		}
		If($Application.SslConnectionEnable)
		{
			line 2 "Enable SSL and TLS protocols: " $Application.SslConnectionEnabled
		}
		If($Application.EncryptionLevel)
		{
			line 2 "Encryption: " $Application.EncryptionLevel
		}
		If($Application.EncryptionRequire)
		{
			line 2 "Minimum requirement: " $Application.EncryptionRequired
		}
	
		line 2 "Start this application without waiting for printers to be created: " -NoNewLine
		If ($Application.WaitOnPrinterCreation) 
		{
			line 0 "False"
		} 
		Else
		{
			line 0 "True"
		}
		
		#appearance properties
		If($Application.WindowType)
		{
			line 2 "Session window size: " $Application.WindowType
		}
		If($Application.ColorDepth)
		{
			line 2 "Maximum color quality: " $Application.ColorDepth
		}
		If($Application.TitleBarHidden)
		{
			line 2 "Hide application title bar: " $Application.TitleBarHidden
		}
		If($Application.MaximizedOnStartup)
		{
			line 2 "Maximize application at startup: " $Application.MaximizedOnStartup
		}
	
	write-output $global:output
	$global:output = $null
	}
}
Else 
{
	line 0 "Application information could not be retrieved"
}

$Applications = $null
$global:output = $null

#servers
$servers = Get-XAServer -EA 0 | sort-object FolderPath, ServerName

If( $? )
{
	line 0 ""
	line 0 "Servers:"
	ForEach($server in $servers)
	{
		line 1 "Name: " $server.ServerName
		line 2 "Product: " $server.CitrixProductName -NoNewLine
		line 0 ", " $server.CitrixEdition -NoNewLine
		line 0 " Edition"
		line 2 "Version: " $server.CitrixVersion
		line 2 "Service Pack: " $server.CitrixServicePack
		line 2 "Operating System Type: " -NoNewLine
		If($server.Is64Bit)
		{
			line 0 "64 bit"
		} 
		Else 
		{
			line 0 "32 bit"
		}
		line 2 "TCP Address: " $server.IPAddresses
		line 2 "Logon: " -NoNewLine
		If($server.LogOnsEnabled)
		{
			line 0 "Enabled"
		} 
		Else 
		{
			line 0 "Disabled"
		}
		line 2 "Product Installation Date: " $server.CitrixInstallDate
		line 2 "Operating System Version: " $server.OSVersion -NoNewLine
		
		#is the server running server 2008?
		If($server.OSVersion.ToString().SubString(0,1) -eq "6")
		{
			$global:Server2008 = $True
		}

		line 0 " " $server.OSServicePack
		line 2 "Zone: " $server.ZoneName
		line 2 "Election Preference: " $server.ElectionPreference
		line 2 "Folder: " $server.FolderPath
		line 2 "Product Installation Path: " $server.CitrixInstallPath
		If($server.ICAPortNumber -gt 0)
		{
			line 2 "ICA Port Number: " $server.ICAPortNumber
		}
		$ServerConfig = Get-XAServerConfiguration -ServerName $Server.ServerName
		If( $? )
		{
			line 2 "Server Configuration Data:"
			If($ServerConfig.AcrUseFarmSettings)
			{
				line 3 "HDX Broadcast\Auto Client Reconnect: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX Broadcast\Auto Client Reconnect: Server is not using farm settings"
				If($ServerConfig.AcrEnabled)
				{
					line 4 "Reconnect automatically"
					line 5 "Log automatic reconnection attempts: " $ServerConfig.AcrLogReconnections 
				}
				Else
				{
					line 4 "Require user authentication"
				}
			}
			line 3 "HDX Broadcast\Browser\Create browser listener on UDP network: " $ServerConfig.BrowserUdpListener
			line 3 "HDX Broadcast\Browser\Server responds to client broadcast messages: " $ServerConfig.BrowserRespondToClientBroadcasts
			If($ServerConfig.DisplayUseFarmSettings)
			{
				line 3 "HDX Broadcast\Display: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX Broadcast\Display: Server is not using farm settings"
				line 4 "Discard queue image\Discard queued image that is replaced by another image: " $ServerConfig.DisplayDiscardQueuedImages
				line 4 "Cache image\Cache image to make scrolling smoother: " $ServerConfig.DisplayCacheImageForSmoothScrolling
				line 4 "Maximum memory to use for each session's graphics (KB): " $ServerConfig.DisplayMaximumGraphicsMemory
				line 4 "Degradation bias: " $ServerConfig.DisplayDegradationBias
				line 4 "Display\Notify user of session degradation: " $ServerConfig.DisplayNotifyUser 
			}
			If($ServerConfig.KeepAliveUseFarmSettings)
			{
				line 3 "HDX Broadcast\Keep-Alive: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX Broadcast\Keep-Alive: Server is not using farm settings"
				line 4 "HDX Broadcast Keep-Alive time-out value: " -NoNewLine
				If($ServerConfig.KeepAliveEnabled)
				{
					line 0 $ServerConfig.KeepAliveTimeout
				}
				Else
				{
					line 0 "Disabled"
				}
			}
			If($ServerConfig.PrinterBandwidth -eq -1)
			{
				line 3 "HDX Broadcast\Printer Bandwidth\Unlimited bandwidth"
			}
			Else
			{
				line 3 "HDX Broadcast\Printer Bandwidth\Limit bandwidth to use (kbps): " $ServerConfig.PrinterBandwidth
			}
			If($ServerConfig.RemoteConsoleUseFarmSettings)
			{
				line 3 "HDX Broadcast\Remote Console Connections: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX Broadcast\Remote Console Connections: Server is not using farm settings"
				line 4 "Remote connections to the console: " $ServerConfig.RemoteConsoleEnabled
			}
			If($ServerConfig.LicenseServerUseFarmSettings)
			{
				line 3 "License Server: Server is using farm settings"
			}
			Else
			{
				line 3 "License Server: Server is not using farm settings"
				line 4 "License server name: " $ServerConfig.LicenseServerName
				line 4 "License server port: " $ServerConfig.LicenseServerPortNumber
			}
			If($ServerConfig.HmrUseFarmSettings)
			{
				line 3 "Health Monitoring & Recovery: Server is using farm settings"
			}
			Else
			{
				line 3 "Health Monitoring & Recovery: Server is not using farm settings"
				line 4 "Run health monitoring tests on this server: " $ServerConfig.HmrEnabled
				If($ServerConfig.HmrEnabled)
				{
					$HMRTests = Get-XAHmrTest -ServerName $Server.ServerName -EA 0
					If( $? )
					{
						line 4 "Health Montoring Tests:"
						ForEach($HMRTest in $HMRTests)
						{
							line 5 "Name            : " $HMRTest.TestName
							line 5 "Interval        : " $HMRTest.Interval
							line 5 "Threshold       : " $HMRTest.Threshold
							line 5 "Time-out        : " $HMRTest.Timeout
							line 5 "Test File Name  : " $HMRTest.FilePath
							line 5 "Recovery Action : " $HMRTest.RecoveryAction
							line 5 "Test Description: " $HMRTest.Description
							line 5 ""
						}
					}
					Else
					{
						line 0 "Health Monitoring & Reporting data could not be retrieved for server " $Server.ServerName
					}
				}
			}
			If($ServerConfig.CpuUseFarmSettings)
			{
				line 3 "CPU Utilization Management: Server is using farm settings"
			}
			Else
			{
				line 3 "CPU Utilization Management: Server is not using farm settings"
				line 4 "CPU Utilization Management: " $ServerConfig.CpuManagementLevel
			}
			If($ServerConfig.MemoryUseFarmSettings)
			{
				line 3 "Memory Optimization: Server is using farm settings"
			}
			Else
			{
				line 3 "Memory Optimization: Server is not using farm settings"
				line 4 "Memory Optimization: " $ServerConfig.MemoryOptimizationEnabled
			}
			If($ServerConfig.ContentRedirectionUseFarmSettings )
			{
				line 3 "HDX Plug and Play/Content Redirection: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX Plug and Play/Content Redirection: Server is not using farm settings"
				line 4 "Content redirection from server to client: " $ServerConfig.ContentRedirectionEnabled
			}
			line 3 "HDX Plug and Play/Shadow Logging/Log shadowing sessions: " $ServerConfig.ShadowLoggingEnabled
			If($ServerConfig.SnmpUseFarmSettings )
			{
				line 3 "SNMP: Server is using farm settings"
			}
			Else
			{
				line 3 "SNMP: Server is not using farm settings"
				# SnmpEnabled is not working
				line 4 "Send session traps to selected SNMP agent on all farm servers: " $ServerConfig.SnmpEnabled
				If( $ServerConfig.SnmpEnabled )
				{
					line 5 "SNMP agent session traps:"
					line 6 "Logon                   : " $ServerConfig.SnmpLogOnEnabled
					line 6 "Logoff                  : " $ServerConfig.SnmpLogOffEnabled
					line 6 "Disconnect              : " $ServerConfig.SnmpDisconnectEnabled
					line 6 "Session limit per server: " $ServerConfig.SnmpLimitEnabled -NoNewLine
					line 0 " " $ServerConfig.SnmpLimitPerServer
				}
			}
			If($ServerConfig.BrowserAccelerationUseFarmSettings )
			{
				line 3 "HDX 3D/Browser Acceleration: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX 3D/Browser Acceleration: Server is not using farm settings"
				line 4 "HDX 3D/Browser Acceleration: " $ServerConfig.BrowserAccelerationEnabled
				If($ServerConfig.BrowserAccelerationEnabled)
				{
					line 5 "Compress JPEG images to improve bandwidth: " $ServerConfig.BrowserAccelerationCompressionEnabled
					If($ServerConfig.BrowserAccelerationCompressionEnabled)
					{
						line 6 "Image compression level: " $ServerConfig.BrowserAccelerationCompressionLevel
						line 6 "Adjust compression level based on available bandwidth: " $ServerConfig.BrowserAccelerationVariableImageCompression
					}
				}
			}
			If($ServerConfig.FlashAccelerationUseFarmSettings )
			{
				line 3 "HDX MediaStream/Flash: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX MediaStream/Flash: Server is not using farm settings"
				line 4 "Enable Flash for XenApp sessions: " $ServerConfig.FlashAccelerationEnabled
				If($ServerConfig.FlashAccelerationEnabled)
				{
					line 5 "Server-side acceleration: " $ServerConfig.FlashAccelerationOption
				}
			}
			If($ServerConfig.MultimediaAccelerationUseFarmSettings )
			{
				line 3 "HDX MediaStream/Multimedia Acceleration: Server is using farm settings"
			}
			Else
			{
				line 3 "HDX MediaStream/Multimedia Acceleration: Server is not using farm settings"
				line 4 "Multimedia acceleration: " $ServerConfig.MultimediaAccelerationEnabled
				If($ServerConfig.MultimediaAccelerationEnabled)
				{
					If($ServerConfig.MultimediaAccelerationDefaultBuffer)
					{
						line 5 "Use the default buffer of 5 seconds"
					}
					Else
					{
						line 5 "Custom buffer in seconds: " $ServerConfig.MultimediaAccelerationCustomBuffer
					}
				}
			}
			line 3 "Virtual IP/Enable virtual IP for this server: " $ServerConfig.VirtualIPEnabled 
			line 3 "Virtual IP/Use farm setting for IP address logging: " $ServerConfig.VirtualIPUseFarmLoggingSettings
			line 3 "Virtual IP/Enable logging of IP address assignment and release on this server: " $ServerConfig.VirtualIPLoggingEnabled
			line 3 "Virtual IP/Enable virtual loopback for this server: " $ServerConfig.VirtualIPLoopbackEnabled
			line 3 "XML Service/Trust requests sent to the XML service: " $ServerConfig.XmlServiceTrustRequests
		}
		Else
		{
			line 0 "Server configuration data could not be retrieved for server " $Server.ServerName
		}
		#applications published to server
		$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			line 2 "Published applications:"
			ForEach($app in $Applications)
			{
				line 0 ""
				line 3 "Display name: " $app.DisplayName
				line 3 "Folder path: " $app.FolderPath
			}
		}
		#Citrix hotfixes installed
		$hotfixes = Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | sort-object HotfixName
		If( $? -and $hotfixes )
		{
			line 0 ""
			line 2 "Citrix Hotfixes:"
			ForEach($hotfix in $hotfixes)
			{
				line 0 ""
				line 3 "Hotfix: " $hotfix.HotfixName
				line 3 "Installed by: " $hotfix.InstalledBy
				line 3 "Installed date: " $hotfix.InstalledOn
				line 3 "Hotfix type: " $hotfix.HotfixType
				line 3 "Valid: " $hotfix.Valid
				line 3 "Hotfixes replaced: "
				ForEach($Replaced in $hotfix.HotfixesReplaced)
				{
					line 4 $Replaced
				}
			}
		}

		line 0 "" 
		write-output $global:output
		$global:output = $null
	}
}
Else 
{
	line 0 "Server information could not be retrieved"
}
$servers = $null
$global:output = $null

$Zones = Get-XAZone -EA 0 | sort-object ZoneName
If( $? )
{
	line 0 ""
	line 0 "Zones:"
	ForEach($Zone in $Zones)
	{
		line 1 "Zone Name: " $Zone.ZoneName
		line 2 "Current Data Collector: " $Zone.DataCollector
		$Servers = Get-XAServer -ZoneName $Zone.ZoneName -EA 0 | sort-object ElectionPreference, ServerName
		If( $? )
		{		
			line 2 "Servers in Zone"
	
			ForEach($Server in $Servers)
			{
				line 3 "Server Name and Preference: " $server.ServerName -NoNewLine
				line 0  " " $server.ElectionPreference
			}
		}
		Else
		{
			line 2 "Unable to enumerate servers in the zone"
		}
		write-output $global:output
		$global:output = $null
	}
}
Else 
{
	line 0 "Zone information could not be retrieved"
}
$Servers = $null
$Zone = $null
$global:output = $null

#Process the nodes in the Advanced Configuration COnsole

#load evaluators
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | sort-object LoadEvaluatorName

If( $? )
{
	line 0 ""
	line 0 "Load Evaluators:"
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		line 1 "Name: " $LoadEvaluator.LoadEvaluatorName
		line 2 "Description: " $LoadEvaluator.Description
		
		If($LoadEvaluator.IsBuiltIn)
		{
			line 2 "Built-in Load Evaluator"
		} 
		Else 
		{
			line 2 "User created load evaluator"
		}
	
		If($LoadEvaluator.ApplicationUserLoadEnabled)
		{
			line 2 "Application User Load Settings"
			line 3 "Report full load when the number of users for this application equals: " $LoadEvaluator.ApplicationUserLoad
			line 3 "Application: " $LoadEvaluator.ApplicationBrowserName
		}
	
		If($LoadEvaluator.ContextSwitchesEnabled)
		{
			line 2 "Context Switches Settings"
			line 3 "Report full load when the number of context switches per second is greater than this value: " $LoadEvaluator.ContextSwitches[1]
			line 3 "Report no load when the number of context switches per second is less than or equal to this value: " $LoadEvaluator.ContextSwitches[0]
		}
	
		If($LoadEvaluator.CpuUtilizationEnabled)
		{
			line 2 "CPU Utilization Settings"
			line 3 "Report full load when the processor utilization percentage is greater than this value: " $LoadEvaluator.CpuUtilization[1]
			line 3 "Report no load when the processor utilization percentage is less than or equal to this value: " $LoadEvaluator.CpuUtilization[0]
		}
	
		If($LoadEvaluator.DiskDataIOEnabled)
		{
			line 2 "Disk Data I/O Settings"
			line 3 "Report full load when the total disk I/O in kilobytes per second is greater than this value: " $LoadEvaluator.DiskDataIO[1]
			line 3 "Report no load when the total disk I/O in kilobytes per second is less than or equal to this value: " $LoadEvaluator.DiskDataIO[0]
		}
	
		If($LoadEvaluator.DiskOperationsEnabled)
		{
			line 2 "Disk Operations Settings"
			line 3 "Report full load when the total number of read and write operations per second is greater than this value: " $LoadEvaluator.DiskOperations[1]
			line 3 "Report no load when the total number of read and write operations per second is less than or equal to this value: " $LoadEvaluator.DiskOperations[0]
		}
	
		If($LoadEvaluator.IPRangesEnabled)
		{
			line 2 "IP Range Settings"
			If($LoadEvaluator.IPRangesAllowed)
			{
				line 3 "Allow " -NoNewLine
			} 
			Else 
			{
				line 3 "Deny " -NoNewLine
			}
			line 3 "client connections from the listed IP Ranges"
			ForEach($IPRange in $LoadEvaluator.IPRanges)
			{
				line 4 "IP Address Ranges: " $IPRange
			}
		}
	
		If($LoadEvaluator.MemoryUsageEnabled)
		{
			line 2 "Memory Usage Settings"
			line 3 "Report full load when the memory usage is greater than this value: " $LoadEvaluator.MemoryUsage[1]
			line 3 "Report no load when the memory usage is less than or equal to this value: " $LoadEvaluator.MemoryUsage[0]
		}
	
		If($LoadEvaluator.PageFaultsEnabled)
		{
			line 2 "Page Faults Settings"
			line 3 "Report full load when the number of page faults per second is greater than this value: " $LoadEvaluator.PageFaults[1]
			line 3 "Report no load when the number of page faults per second is less than or equal to this value: " $LoadEvaluator.PageFaults[0]
		}
	
		If($LoadEvaluator.PageSwapsEnabled)
		{
			line 2 "Page Swaps Settings"
			line 3 "Report full load when the number of page swaps per second is greater than this value: " $LoadEvaluator.PageSwaps[1]
			line 3 "Report no load when the number of page swaps per second is less than or equal to this value: " $LoadEvaluator.PageSwaps[0]
		}
	
		If($LoadEvaluator.ScheduleEnabled)
		{
			line 2 "Scheduling Settings"
			line 3 "Sunday Schedule: " $LoadEvaluator.SundaySchedule
			line 3 "Monday Schedule: " $LoadEvaluator.MondaySchedule
			line 3 "Tuesday Schedule: " $LoadEvaluator.TuesdaySchedule
			line 3 "Wednesday Schedule: " $LoadEvaluator.WednesdaySchedule
			line 3 "Thursday Schedule: " $LoadEvaluator.ThursdaySchedule
			line 3 "Friday Schedule: " $LoadEvaluator.FridaySchedule
			line 3 "Saturday Schedule: " $LoadEvaluator.SaturdaySchedule
		}
	
		line 0 ""
		write-output $global:output
		$global:output = $null
	}
		
}
Else 
{
	line 0 "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $null
$global:output = $null

$Policies = Get-XAPolicy -EA 0 | sort-object PolicyName
If( $? -and $Policies)
{
	line 0 ""
	line 0 "Policies:"
	ForEach($Policy in $Policies)
	{
		line 1 "Policy Name: " $Policy.PolicyName
		line 2 "Description: " $Policy.Description
		line 2 "Enabled: " $Policy.Enabled
		line 2 "Priority: " $Policy.Priority

		$filter = Get-XAPolicyFilter -PolicyName $Policy.PolicyName -EA 0

		If( $? )
		{
			If($Filter)
			{
				line 2 "Policy Filters:"
				
				If($Filter.AccessControlEnabled)
				{
					If($Filter.AllowConnectionsThroughAccessGateway)
					{
						line 3 "Apply to connections made through Access Gateway"
						If($Filter.AccessSessionConditions)
						{
							line 4 "Any connection that meets any of the following filters"
							ForEach($Condition in $Filter.AccessSessionConditions)
							{
								$Colon = $Condition.IndexOf(":")
								$CondName = $Condition.SubString(0,$Colon)
								$CondFilter = $Condition.SubString($Colon+1)
								Line 5 "Access Gateway Farm Name: " $CondName -NoNewLine
								Line 0 " Filter Name: " $CondFilter
							}
						}
						Else
						{
							line 4 "Any connection"
						}
						line 4 "Apply to all other connections: " $Filter.AllowOtherConnections
					}
					Else
					{
						line 3 "Do not apply to connections made through Access Gateway"
					}
				}
				If($Filter.ClientIPAddressEnabled)
				{
					line 3 "Apply to all client IP addresses: " $Filter.ApplyToAllClientIPAddresses
					If($Filter.AllowedIPAddresses)
					{
						line 3 "Allowed IP Addresses:"
						ForEach($Allowed in $Filter.AllowedIPAddresses)
						{
							line 4 $Allowed
						}
					}
					If($Filter.DeniedIPAddresses)
					{
						line 3 "Denied IP Addresses:"
						ForEach($Denied in $Filter.DeniedIPAddresses)
						{
							line 4 $Denied
						}
					}
				}
				If($Filter.ClientNameEnabled)
				{
					line 3 "Apply to all client names: " $Filter.ApplyToAllClientNames
					If($Filter.AllowedClientNames)
					{
						line 3 "Allowed Client Names:"
						ForEach($Allowed in $Filter.AllowedClientNames)
						{
							line 4 $Allowed
						}
					}
					If($Filter.DeniedClientNames)
					{
						line 3 "Denied Client Names:"
						ForEach($Denied in $Filter.DeniedClientNames)
						{
							line 4 $Denied
						}
					}
				}
				If($Filter.ServerEnabled)
				{
					If($Filter.AllowedServerNames)
					{
						line 3 "Allowed Server Names:"
						ForEach($Allowed in $Filter.AllowedServerNames)
						{
							line 4 $Allowed
						}
					}
					If($Filter.DeniedServerNames)
					{
						line 3 "Denied Server Names:"
						ForEach($Denied in $Filter.DeniedServerNames)
						{
							line 4 $Denied
						}
					}
					If($Filter.AllowedServerFolders)
					{
						line 3 "Allowed Server Folders:"
						ForEach($Allowed in $Filter.AllowedServerFolders)
						{
							line 4 $Allowed
						}
					}
					If($Filter.DeniedServerFolders)
					{
						line 3 "Denied Server Folders:"
						ForEach($Denied in $Filter.DeniedServerFolders)
						{
							line 4 $Denied
						}
					}
				}
				If($Filter.AccountEnabled)
				{
					line 3 "Apply to all explicit (non-anonymous) users: " $Filter.ApplyToAllExplicitAccounts
					line 3 "Apply to anonymous users: " $Filter.ApplyToAnonymousAccounts
					If($Filter.AllowedAccounts)
					{
						line 3 "Allowed Accounts:"
						ForEach($Allowed in $Filter.AllowedAccounts)
						{
							line 4 $Allowed
						}
					}
					If($Filter.DeniedAccounts)
					{
						line 3 "Denied Accounts:"
						ForEach($Denied in $Filter.DeniedAccounts)
						{
							line 4 $Denied
						}
					}
				}
			}
			Else
			{
				line 2 "No filter information"
			}
		}
		Else
		{
			Line 2 "Unable to retrieve Filter settings"
		}

		$Settings = Get-XAPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0

		If( $? )
		{
			line 2 "Policy Settings:"
			ForEach($Setting in $Settings)
			{
				#HDX 3D
				If($Setting.ImageAccelerationState -ne "NotConfigured")
				{
					line 3 "HDX 3D\Progressive Display\Progressive Display: " $Setting.ImageAccelerationState
					If($Setting.ImageAccelerationState -eq "Enabled")
					{
						line 3 "HDX 3D\Progressive Display\Progressive Display\Compression level: " $Setting.ImageAccelerationCompressionLevel
						line 3 "HDX 3D\Progressive Display\Progressive Display\Compression level\Restrict compression to connections under this bandwidth: " $Setting.ImageAccelerationCompressionIsRestricted
						If($Setting.ImageAccelerationCompressionIsRestricted)
						{
							line 3 "HDX 3D\Progressive Display\Progressive Display\Compression level\Restrict compression to connections under this bandwidth\Threshhold (Kb/sec): " $Setting.ImageAccelerationCompressionLimit	
						}
						line 3 "HDX 3D\Progressive Display\Progressive Display\SpeedScreen Progressive Display compression level: " $Setting.ImageAccelerationProgressiveLevel
						line 3 "HDX 3D\Progressive Display\Progressive Display\Restrict compression to connections under this bandwidth: " $Setting.ImageAccelerationProgressiveIsRestricted
						If($Setting.ImageAccelerationProgressiveIsRestricted)
						{
							line 3 "HDX 3D\Progressive Display\Progressive Display\Restrict compression to connections under this bandwidth\Threshhold (Kb/sec): " $Setting.ImageAccelerationProgressiveLimit	
						}
						line 3 "HDX 3D\Progressive Display\Progressive Display\Use Heavyweight compression (extra CPU, retains quality): " $Setting.ImageAccelerationIsHeavyweightUsed
					}
				}
				
				#HDX Broadcast
				If($Setting.TurnWallpaperOffState -ne "NotConfigured")
				{
					line 3 "HDX Broadcast\Visual Effects\Turn off desktop wallpaper: " $Setting.TurnWallpaperOffState
					If($Setting.TurnWallpaperOffState -eq "Enabled")
					{
						line 3 "HDX Broadcast\Visual Effects\Turn off desktop wallpaper\Turn Off Desktop Wallpaper"
					}
				}
					
				#menu animation 2008 only
				If($global:Server2008)
				{
					If($Setting.TurnMenuAnimationsOffState -ne "NotConfigured")
					{
						line 3 "HDX Broadcast\Visual Effects\Turn off menu animations: " $Setting.TurnMenuAnimationsOffState
						If($Setting.TurnMenuAnimationsOffState -eq "Enabled")
						{
							line 3 "HDX Broadcast\Visual Effects\Turn off menu animations\Turn Off Menu and Window Animations"
						}
					}
				}
				If($Setting.TurnWindowContentsOffState -ne "NotConfigured")
				{
					line 3 "HDX Broadcast\Visual Effects\Turn off window contents while dragging: " $Setting.TurnWindowContentsOffState
					If($Setting.TurnWindowContentsOffState -eq "Enabled")
					{
						line 3 "HDX Broadcast\Visual Effects\Turn off window contents while dragging\Turn Off Windows Contents While Dragging"
					}
				}
				If($Setting.SessionAudioState -ne "NotConfigured")
				{
					line 3 "Session Limits\Audio: " $Setting.SessionAudioState
					If($Setting.SessionAudioState -eq "Enabled")
					{
						line 3 "Session Limits\Audio\Limit (Kb/sec): " $Setting.SessionAudioLimit
					}
				}
				If($Setting.SessionClipboardState -ne "NotConfigured")
				{
					line 3 "Session Limits\Clipboard: " $Setting.SessionClipboardState
					If($Setting.SessionClipboardState -eq "Enabled")
					{
						line 3 "Session Limits\Clipboard\Limit (Kb/sec): " $Setting.SessionClipboardLimit
					}
				}
				If($Setting.SessionComportsState -ne "NotConfigured")
				{
					line 3 "Session Limits\COM Ports: " $Setting.SessionComportsState
					If($Setting.SessionComportsState -eq "Enabled")
					{
						line 3 "Session Limits\COM Ports\Limit (Kb/sec): " $Setting.SessionComportsLimit
					}
				}
				If($Setting.SessionDrivesState -ne "NotConfigured")
				{
					line 3 "Session Limits\Drives: " $Setting.SessionDrivesState
					If($Setting.SessionDrivesState -eq "Enabled")
					{
						line 3 "Session Limits\Drives\Limit (Kb/sec): " $Setting.SessionDrivesLimit
					}
				}
				If($Setting.SessionLptPortsState -ne "NotConfigured")
				{
					line 3 "Session Limits\LPT Ports: " $Setting.SessionLptPortsState
					If($Setting.SessionLptPortsState -eq "Enabled")
					{
						line 3 "Session Limits\LPT Ports\Limit (Kb/sec): " $Setting.SessionLptPortsLimit
					}
				}
				If($Setting.SessionOemChannelsState -ne "NotConfigured")
				{
					line 3 "Session Limits\OEM Virtual Channels: " $Setting.SessionOemChannelsState
					If($Setting.SessionOemChannelsState -eq "Enabled")
					{
						line 3 "Session Limits\OEM Virtual Channels\Limit (Kb/sec): " $Setting.SessionOemChannelsLimit
					}
				}
				If($Setting.SessionOverallState -ne "NotConfigured")
				{
					line 3 "Session Limits\Overall Session: " $Setting.SessionOverallState
					If($Setting.SessionOverallState -eq "Enabled")
					{
						line 3 "Session Limits\Overall Session\Limit (Kb/sec): " $Setting.SessionOverallLimit
					}
				}
				If($Setting.SessionPrinterBandwidthState -ne "NotConfigured")
				{
					line 3 "Session Limits\Printer: " $Setting.SessionPrinterBandwidthState
					If($Setting.SessionPrinterBandwidthState -eq "Enabled")
					{
						line 3 "Session Limits\Printer\Limit (Kb/sec): " $Setting.SessionPrinterBandwidthLimit
					}
				}
				If($Setting.SessionTwainRedirectionState -ne "NotConfigured")
				{
					line 3 "Session Limits\TWAIN Redirection: " $Setting.SessionTwainRedirectionState
					If($Setting.SessionTwainRedirectionState -eq "Enabled")
					{
						line 3 "Session Limits\TWAIN Redirection\Limit (Kb/sec): " $Setting.SessionTwainRedirectionLimit
					}
				}

				#Session Limits % Server 2008 only
				If($global:Server2008)
				{
					If($Setting.SessionAudioPercentState -ne "NotConfigured")
					{
						line 3 'Session Limits (%)\Audio: ' $Setting.SessionAudioPercentState
						If($Setting.SessionAudioPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\Audio\Limit (%): ' $Setting.SessionAudioPercentLimit
						}
					}
					If($Setting.SessionClipboardPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\Clipboard: " $Setting.SessionClipboardPercentState
						If($Setting.SessionClipboardPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\Clipboard\Limit (%): ' $Setting.SessionClipboardPercentLimit
						}
					}
					If($Setting.SessionComportsPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\COM Ports: " $Setting.SessionComportsPercentState
						If($Setting.SessionComportsPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\COM Ports\Limit (%): ' $Setting.SessionComportsPercentLimit
						}
					}
					If($Setting.SessionDrivesPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\Drives: " $Setting.SessionDrivesPercentState
						If($Setting.SessionDrivesPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\Drives\Limit (%): ' $Setting.SessionDrivesPercentLimit
						}
					}
					If($Setting.SessionLptPortsPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\LPT Ports: " $Setting.SessionLptPortsPercentState
						If($Setting.SessionLptPortsPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\LPT Ports\Limit (%): ' $Setting.SessionLptPortsPercentLimit
						}
					}
					If($Setting.SessionOemChannelsPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\OEM Virtual Channels: " $Setting.SessionOemChannelsPercentState
						If($Setting.SessionOemChannelsPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\OEM Virtual Channels\Limit (%): ' $Setting.SessionOemChannelsPercentLimit
						}
					}
					If($Setting.SessionPrinterPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\Printer: " $Setting.SessionPrinterPercentState
						If($Setting.SessionPrinterPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\Printer\Limit (%): ' $Setting.SessionPrinterPercentLimit
						}
					}
					If($Setting.SessionTwainRedirectionPercentState -ne "NotConfigured")
					{
						line 3 "Session Limits (%)\TWAIN Redirection: " $Setting.SessionTwainRedirectionPercentState
						If($Setting.SessionTwainRedirectionPercentState -eq "Enabled")
						{
							line 3 'Session Limits (%)\TWAIN Redirection\Limit (%): ' $Setting.SessionTwainRedirectionPercentLimit
						}
					}
				}

				#HDX Plug-n-Play
				If($Setting.ClientMicrophonesState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Audio\Microphones: " $Setting.ClientMicrophonesState
					If($Setting.ClientMicrophonesState -eq "Enabled")
					{
						If($Setting.ClientMicrophonesAreUsed)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Audio\Microphones\Use client microphones for audio input"
						}
						Else
						{
							line 3 "HDX Plug-n-Play\Client Resources\Audio\Microphones\Do not use client microphones for audio input"
						}
					}
				}
				If($Setting.ClientSoundQualityState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Sound quality: " $Setting.ClientSoundQualityState
					If($Setting.ClientSoundQualityState)
					{
						line 3 "HDX Plug-n-Play\Client Resources\Sound quality\Maximum allowable client audio quality: " $Setting.ClientSoundQualityLevel
					}
				}
				If($Setting.TurnClientAudioMappingOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Turn off speakers: " $Setting.TurnClientAudioMappingOffState
					If($Setting.TurnClientAudioMappingOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\Turn off speakers\Turn off audio mapping to client speakers"
					}
				}
				If($Setting.ClientDrivesState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Drives\Connection: " $Setting.ClientDrivesState
					If($Setting.ClientDrivesState -eq "Enabled")
					{
						If($Setting.ClientDrivesAreConnected)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Connection\Connect Client Drives at Logon"
						}
						Else
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Connection\Do Not Connect Client Drives at Logon"
						}
					}
				}
				If($Setting.ClientDriveMappingState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Drives\Mappings: " $Setting.ClientDriveMappingState
					If($Setting.ClientDriveMappingState -eq "Enabled")
					{
						If($Setting.TurnFloppyDriveMappingOff)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Mappings\Turn off Floppy disk drives"	
						}
						If($Setting.TurnHardDriveMappingOff)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Mappings\Turn off Hard drives"	
						}
						If($Setting.TurnCDRomDriveMappingOff)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Mappings\Turn off CD-ROM drives"	
						}
						If($Setting.TurnRemoteDriveMappingOff)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Mappings\Turn off Remote drives"	
						}
					}
				}
				If($Setting.ClientAsynchronousWritesState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Drives\Optimize\Asynchronous writes: " $Settings.ClientAsynchronousWritesState
					If($Settings.ClientAsynchronousWritesState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\Drives\Optimize\Asynchronous writes\Turn on asynchronous disk writes to client disks"
					}
	
					If($global:Server2008)
					{
						line 3 "HDX Plug-n-Play\Client Resources\Drives\Special folder redirection: " $Setting.TurnSpecialFolderRedirectionOffState
						If($Setting.TurnSpecialFolderRedirectionOffState -eq "Enabled")
						{
							line 3 "HDX Plug-n-Play\Client Resources\Drives\Special folder redirection\Do not allow special folder redirection"
						}
					}
				}

				If($Setting.TurnComPortsOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Ports\Turn off COM ports: " $Setting.TurnComPortsOffState
					If($Setting.TurnComPortsOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\Ports\Turn off COM ports\Turn off client COM ports"
					}
				}
				If($Setting.TurnLptPortsOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Ports\Turn off LPT ports: " $Setting.TurnLptPortsOffState
					If($Setting.TurnLptPortsOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\Ports\Turn off LPT ports\Turn off client LPT ports"
					}
				}
				If($Setting.TurnVirtualComPortMappingOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\PDA Devices\Turn on automatic virtual COM port mapping: " $Setting.TurnVirtualComPortMappingOffState
					If($Setting.TurnVirtualComPortMappingOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\PDA Devices\Turn on automatic virtual COM port mapping\Turn on virtual COM port mapping"
					}
				}
				If($Setting.TwainRedirectionState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Other\Configure TWAIN redirection: " $Setting.TwainRedirectionState
					If($Setting.TwainRedirectionState -eq "Enabled")
					{
						If($Setting.TwainRedirectionAllowed)
						{
							line 3 "HDX Plug-n-Play\Client Resources\Other\Configure TWAIN redirection\Allow TWAIN redirection"
							If($Setting.TwainRedirectionImageCompression -eq "NoCompression")
							{
								line 3 "HDX Plug-n-Play\Client Resources\Other\Configure TWAIN redirection\Allow TWAIN redirection\Do not use lossy compression for high color images"
							}
							Else
							{
								line 3 "HDX Plug-n-Play\Client Resources\Other\Configure TWAIN redirection\Allow TWAIN redirection\Use lossy compression for high color images: " $Setting.TwainRedirectionImageCompression
							}
						}
						Else
						{
							line 3 "HDX Plug-n-Play\Client Resources\Other\Configure TWAIN redirection\Allow TWAIN redirection\Do not allow TWAIN redirection"
						}
					}
				}
				If($Setting.TurnClipboardMappingOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Other\Turn off clipboard mapping: " $Setting.TurnClipboardMappingOffState
					If($Setting.TurnClipboardMappingOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\Other\Turn off clipboard mapping\Turn Off Client Clipboard Mapping"
					}
				}
				If($Setting.TurnOemVirtualChannelsOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Resources\Other\Turn off OEM virtual channels: " $Setting.TurnOemVirtualChannelsOffState
					If($Setting.TurnOemVirtualChannelsOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Resources\Other\Turn off OEM virtual channels\Turn Off OEM Virtual Channels"
					}
				}

				If($Setting.TurnAutoClientUpdateOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Client Maintenance\Turn off auto client update: " $Setting.TurnAutoClientUpdateOffState
					If($Setting.TurnAutoClientUpdateOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Client Maintenance\Turn off auto client update\Turn Off Auto Client Update"
					}
				}
				If($Setting.ClientPrinterAutoCreationState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Client Printers\Auto-creation: " $Setting.ClientPrinterAutoCreationState
					If($Setting.ClientPrinterAutoCreationState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Printing\Client Printers\Auto-creation\When connecting " $Setting.ClientPrinterAutoCreationOption
					}
				}

				If($Setting.LegacyClientPrintersState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Client Printers\Legacy client printers: " $Setting.LegacyClientPrintersState
					If($Setting.LegacyClientPrintersState -eq "Enabled")
					{
						If($Setting.LegacyClientPrintersDynamic)
						{
							line 3 "HDX Plug-n-Play\Printing\Client Printers\Legacy client printers\Create dynamic session-private client printers"
						}
						Else
						{
							line 3 "HDX Plug-n-Play\Printing\Client Printers\Legacy client printers\Create old-style client printers"
						}
					}
				}
				If($Setting.PrinterPropertiesRetentionState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Client Printers\Printer properties retention: " $Setting.PrinterPropertiesRetentionState
					If($Setting.PrinterPropertiesRetentionState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Printing\Client Printers\Printer properties retention\Printer properties should be " $Setting.PrinterPropertiesRetentionOption
					}
				}
				If($Setting.PrinterJobRoutingState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Client Printers\Print job routing: " $Setting.PrinterJobRoutingState
					If($Setting.PrinterJobRoutingState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Printing\Client Printers\Print job routing\For client printers on a network printer server: " -NoNewLine
						If($Setting.PrinterJobRoutingDirect)
						{
							line 0 "Connect directly to network print server if possible"
						}
						Else
						{
							line 0 "Always connect indirectly as a client printer"
						}
					}
				}
				If($Setting.TurnClientPrinterMappingOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Client Printers\Turn off client printer mapping: " $Setting.TurnClientPrinterMappingOffState
					If($Setting.TurnClientPrinterMappingOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Printing\Client Printers\Turn off client printer mapping\Turn off client printer mapping"
					}
				}
				If($Setting.DriverAutoInstallState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Drivers\Native printer driver auto-install: " $Setting.DriverAutoInstallState
					If($Setting.DriverAutoInstallState -eq "Enabled")
					{
						If($Setting.DriverAutoInstallAsNeeded)
						{
							line 3 "HDX Plug-n-Play\Printing\Drivers\Native printer driver auto-install\Install Windows native drivers as needed"
						}
						Else
						{
							line 3 "HDX Plug-n-Play\Printing\Drivers\Native printer driver auto-install\Do not automatically install drivers"
						}
					}
				}
				If($Setting.UniversalDriverState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Drivers\Universal driver: " $Setting.UniversalDriverState
					If($Setting.UniversalDriverState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Printing\Drivers\Universal driver\When auto-creating client printers: " $Setting.UniversalDriverOption
					}
				}
				If($Setting.SessionPrintersState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Printing\Session printers\Session printers: " $Setting.SessionPrintersState
					If($Setting.SessionPrintersState -eq "Enabled")
					{
						If($Setting.SessionPrinterList)
						{
							line 3 "HDX Plug-n-Play\Printing\Session printers\Session printers\Network printers to connect at logon:"
							ForEach($Printer in $Setting.SessionPrinterList)
							{
								Line 7 $Printer
							}
						}
						line 3 "HDX Plug-n-Play\Printing\Session printers\Client's default printer: " -NoNewLine
						If($Setting.SessionPrinterDefaultOption -eq "SetToPrinterIndex")
						{
							line 0 $Setting.SessionPrinterList[$Setting.SessionPrinterDefaultIndex]
						}
						Else
						{
							line 0 $Setting.SessionPrinterDefaultOption
						}
					}
				}
				If($Setting.ContentRedirectionState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Content Redirection\Server to client: " $Setting.ContentRedirectionState
					If($Setting.ContentRedirectionState -eq "Enabled")
					{
						If($Setting.ContentRedirectionIsUsed)
						{
							line 3 "HDX Plug-n-Play\Content Redirection\Server to client\Use Content Redirection from server to client"
						}
						Else
						{
							line 3 "HDX Plug-n-Play\Content Redirection\Server to client\Do not use Content Redirection from server to client"
						}
					}
				}
				If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Time Zones\Do not estimate local time for legacy clients: " $Setting.TurnClientLocalTimeEstimationOffState
					If($Setting.TurnClientLocalTimeEstimationOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Time Zones\Do not estimate local time for legacy clients\Do Not Estimate Client Local Time"
					}
				}
				If($Setting.TurnClientLocalTimeEstimationOffState -ne "NotConfigured")
				{
					line 3 "HDX Plug-n-Play\Time Zones\Do not use Client's local time: " $Setting.TurnClientLocalTimeOffState
					If($Setting.TurnClientLocalTimeOffState -eq "Enabled")
					{
						line 3 "HDX Plug-n-Play\Time Zones\Do not use Client's local time\Do Not Use Client's Local Time"
					}
				}

				#User Workspace
				If($Setting.ConcurrentSessionsState -ne "NotConfigured")
				{
					line 3 "User Workspace\Connections\Limit total concurrent sessions: " $Setting.ConcurrentSessionsState
					If($Setting.ConcurrentSessionsState -eq "Enabled")
					{
						line 3 "User Workspace\Connections\Limit total concurrent sessions\Limit: " $Setting.ConcurrentSessionsLimit
					}
				}
				If($Setting.ZonePreferenceAndFailoverState -ne "NotConfigured")
				{
					line 3 "User Workspace\Connections\Zone preference and failover: " $Setting.ZonePreferenceAndFailoverState
					If($Setting.ZonePreferenceAndFailoverState -eq "Enabled")
					{
						line 3 "User Workspace\Connections\Zone preference and failover\Zone preference settings:"
						ForEach($Pref in $Setting.ZonePreferences)
						{
							line 7 $Pref
						}
					}
				}
				If($Setting.ShadowingState -ne "NotConfigured")
				{
					line 3 "User Workspace\Shadowing\Configuration: " $Setting.ShadowingState
					If($Setting.ShadowingState -eq "Enabled")
					{
						If($Setting.ShadowingAllowed)
						{
							line 3 "User Workspace\Shadowing\Configuration\Allow Shadowing"
							line 3 "User Workspace\Shadowing\Configuration\Prohibit Being Shadowed Without Notification: " $Setting.ShadowingProhibitedWithoutNotification
							line 3 "User Workspace\Shadowing\Configuration\Prohibit Remote Input When Being Shadowed: " $Setting.ShadowingRemoteInputProhibited
						}
						Else
						{
							line 3 "User Workspace\Shadowing\Configuration\Do Not Allow Shadowing"
						}
					}
				}
				If($Setting.ShadowingPermissionsState -ne "NotConfigured")
				{
					line 3 "User Workspace\Shadowing\Permissions: " $Setting.ShadowingPermissionsState
					If($Setting.ShadowingPermissionsState -eq "Enabled")
					{
						If($Setting.ShadowingAccountsAllowed)
						{
							line 3 "User Workspace\Shadowing\Permissions\Accounts allowed to shadow:"
							ForEach($Allowed in $Setting.ShadowingAccountsAllowed)
							{
								line 7 $Allowed
							}
						}
						If($Setting.ShadowingAccountsDenied)
						{
							line 3 "User Workspace\Shadowing\Permissions\Accounts denied from shadowing:"
							ForEach($Denied in $Setting.ShadowingAccountsDenied)
							{
								line 7 $Denied
							}
						}
					}
				}
				If($Setting.CentralCredentialStoreState -ne "NotConfigured")
				{
					line 3 "User Workspace\Single Sign-On\Central Credential Store: " $Setting.CentralCredentialStoreState
					If($Setting.CentralCredentialStoreState -eq "Enabled")
					{
						If($Setting.CentralCredentialStorePath)
						{
							line 3 "User Workspace\Single Sign-On\Central Credential Store\UNC path of Central Credential Store: " $Setting.CentralCredentialStorePath
						}
						Else
						{
							line 3 "User Workspace\Single Sign-On\Central Credential Store\No UNC path to Central Credential Store entered"
						}
					}
				}
				If($Setting.TurnPasswordManagerOffState -ne "NotConfigured")
				{
					line 3 "User Workspace\Single Sign-On\Do not use Citrix Password Manager: " $Setting.TurnPasswordManagerOffState
					If($Setting.TurnPasswordManagerOffState -eq "Enabled")
					{
						line 3 "User Workspace\Single Sign-On\Do not use Citrix Password Manager\Do not use Citrix Password Manager"
					}
				}
				If($Setting.StreamingDeliveryProtocolState -ne "NotConfigured")
				{
					line 3 "User Workspace\Streamed Applications\Configure delivery protocol: " $Setting.StreamingDeliveryProtocolState
					If($Setting.StreamingDeliveryProtocolState -eq "Enabled")
					{
						line 3 "User Workspace\Streamed Applications\Configure delivery protocol\Streaming Delivery Protocol option: " $Setting.StreamingDeliveryOption
					}
				}

				#Security
				If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
				{
					line 3 "Security\Encryption\SecureICA encryption: " $Setting.SecureIcaEncriptionState
					If($Setting.SecureIcaEncriptionState -eq "Enabled")
					{
						line 3 "Security\Encryption\SecureICA encryption\Encryption level: " $Setting.SecureIcaEncriptionLevel
					}
				}
				
				#Service Level (2008 only)
				If($global:Server2008)
				{
					If($Setting.SecureIcaEncriptionState -ne "NotConfigured")
					{
						line 3 "Service Level\Session Importance: " $Setting.SessionImportanceState
						If($Setting.SessionImportanceState -eq "Enabled")
						{
							line 3 "Service Level\Session Importance\Importance level: " $Setting.SessionImportanceLevel
						}
					}
				}
			}
		}
		Else
		{
			Line 2 "Unable to retrieve settings"
		}
	
		write-output $global:output
		$global:output = $null
		$Settings = $null
		$Filter = $null
	}
}
Else 
{
	line 0 "Citrix Policy information could not be retrieved."
}
$Policies = $null
$global:output = $null

#printer drivers
$PrinterDrivers = Get-XAPrinterDriver -EA 0 | sort-object DriverName

If( $? -and $PrinterDrivers)
{
	line 0 ""
	line 0 "Print Drivers:"
	ForEach($PrinterDriver in $PrinterDrivers)
	{
		line 1 "Driver  : " $PrinterDriver.DriverName
		line 1 "Platform: " $PrinterDriver.OSVersion
		line 1 "64 bit? : " $PrinterDriver.Is64Bit
		line 0 ""
	}
	write-output $global:output
	$global:output = $null
}
Else 
{
	line 0 "Printer driver information could not be retrieved"
}

$PrintDrivers = $null
$global:output = $null

#printer driver mappings
$PrinterDriverMappings = Get-XAPrinterDriverMapping -EA 0 | sort-object ClientDriverName

If( $? -and $PrinterDriverMappings)
{
	line 0 ""
	line 0 "Print Driver Mappings:"
	ForEach($PrinterDriverMapping in $PrinterDriverMappings)
	{
		line 1 "Client Driver: " $PrinterDriverMapping.ClientDriverName
		line 1 "Server Driver: " $PrinterDriverMapping.ServerDriverName
		line 1 "Platform: " $PrintDriverMapping.OSVersion
		line 1 "64 bit? : " $PrinterDriverMapping.Is64Bit
		line 0 ""
	}
	write-output $global:output
	$global:output = $null
}
Else 
{
	line 0 "Printer driver mapping information could not be retrieved"
}

$PrintDriverMappings = $null
$global:output = $null

If( $Global:ConfigLog)
{
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
			line 0 ""
			line 0 "Configuration Log Report:"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				line 0 ""
				Line 1 "Date: " $ConfigLogItem.Date
				Line 1 "Account: " $ConfigLogItem.Account
				Line 1 "Change description: " $ConfigLogItem.Description
				Line 1 "Type of change: " $ConfigLogItem.TaskType
				Line 1 "Type of item: " $ConfigLogItem.ItemType
				Line 1 "Name of item: " $ConfigLogItem.ItemName
			}
			Write-Output $global:output
			$global:output = $null
		} 
		Else 
		{
			line 0 "Configuration log report could not be retrieved"
		}
		Write-Output $global:output
		$ConfigLogReport = $null
		$global:output = $null
	}
	Else 
	{
		line 0 "XA5ConfigLog.udl file was not found"
	}
}