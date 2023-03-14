Clear-Host
Write-Host "Start to Check the Exchange Server Healthy...`r`n"



#Suppress Output if caught the errors
$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

#snapin
$isPSSnapinLoaded = Get-PSSnapin -name "Microsoft.Exchange.Management.PowerShell.E2010"

if ($isPSSnapinLoaded -eq $NULL)
{
	Add-PSSnapin -Name "Microsoft.Exchange.Management.PowerShell.E2010"
} else
{
	break
}

.'C:\Program Files\Microsoft\Exchange Server\V15\Bin\RemoteExchange.ps1'

Connect-ExchangeServer -auto



#parameters
$CommonServices = @("W32Time", "Winmgmt", "Dnscache", "RpcEptMapper", "rpcss")
$CheckTime = Get-Date -Format "yyyy-MM-dd.HH.mm"
$CheckTime1 = Get-Date -Format "yyyy-MM-dd HH:mm"
$WarningLevel = 20.0;
$WorkFolder = "c:\xiaomi"
#Define the output for each check object
$DiskCheckResult = @()
$DiskCheckError = @()
$TimeDiffCheckResult = @()
$TimeDiffCheckError = @()
$ServiceCheckResult = @()
$ServiceCheckError = @()
$DBCopyCheckResult = @()
$DBCopyCheckError = @()
$DBBackupCheckResult = @()
$DBBackupCheckError = @()
$QueueCheckResult = @()
$QueueCheckError = @()
$EventLogCheckResult = @()
$EventLogCheckError = @()
$EWSTestResult = @()
$EWSTestError = @()
$OWATestResult = @()
$OWATestError = @()
$ECPTestResult = @()
$ECPTestError = @()
$ASTestResult = @()
$ASTestError = @()
$MAPITestResult = @()
$MAPITestError = @()
$REPLTestResult = @()
$REPLTestError = @()
$DBPeocountCheckResult = @()
$DBPeocountCheckError = @()
$ServerComponentStateResult = @()

#serverlist
$ExchangeServers = get-exchangeserver
$CASServers = Get-ClientAccessServer  | where-object {$_.name -like "*cas*"}
$MBXServers = Get-MailboxServer | where-object {$_.name -like "*box*"}
$DBs =  Get-MailboxDatabase -Status | where-object {$_.server -like "*box*"}
$HTServers = Get-TransportService 
$MBXDBs = Get-MailboxDatabase    



# Start to check time differentiation between Exchange server and DC.
Write-Host "Checking Time differential"  `r
foreach($Server in $ExchangeServers)
{
	## Start to check time differentiation between Exchange server and DC.
	#Write-Host "Checking Time differential"  `r
	
	#$ExServerTime = net time  \\$Server
	#$ExServerTime, $tmp, $Success = $ExServerTime
	#$ExServerTime = ($ExServerTime.split("是")[-1]).trim()
	#$ExServerTime = [datetime]::ParseExact($ExServerTime, "M/d/yyyy h:mm:ss tt", $null)
	
	#$DCTime = net time
	#$DCTime,$tmp,$Success = $DCTime
	#$DCTime = ($DCTime.split("是")[-1]).trim()
	#$DCTime = [datetime]::ParseExact($DCTime, "M/d/yyyy h:mm:ss tt", $null)
	
	#$TimeDiff = ($DCTime - $ExServerTime).TotalSeconds
	#$TimeDiff = "{0:N0}" -f $TimeDiff
	$DCTime = icm mioffice-dc1 {Get-Date}
    $ExServerTime = icm $Server {Get-Date}
    $TimeDiff = ($DCTime - $ExServerTime).totalseconds
	$TimeDiffObject = New-Object PSObject
	Add-Member -InputObject $TimeDiffObject NoteProperty "Server Name" $Server.name
	Add-Member -InputObject $TimeDiffObject NoteProperty "Server Time" $ExServerTime
	Add-Member -InputObject $TimeDiffObject NoteProperty "DC Time" $DCTime
	Add-Member -InputObject $TimeDiffObject NoteProperty "Time Diff" $TimeDiff

	$TimeDiffCheckResult = $TimeDiffCheckResult += $TimeDiffObject

}

## End of code for time checking.
	
	##Check Disk free space...
	Write-Host "Checking Disk free space..."  `r
	foreach($Server in $ExchangeServers)
	{
		$disks = Get-WMIObject -ComputerName $Server -Class WIN32_Volume
			foreach($disk in $disks)
		{
			$CheckTime1 = Get-Date -Format "yyyy-MM-dd HH:mm"
			if(($disk.name.contains("\\") -ne "true") -and (([Math]::Round(($disk.Capacity /1048576)) -ne 0)))
			{
				$Capacity = [Math]::Round(($disk.Capacity /1048576), 0)
				$Freespace = [Math]::Round(($disk.Freespace / 1048576), 0) 
				$PercentAvailable = [Math]::Round((100.0 * $disk.Freespace / $disk.Capacity), 1) 
				$DiskObject = new-Object PSObject
				Add-Member -InputObject $DiskObject NoteProperty "Check Date"	$CheckTime1
				Add-Member -InputObject $DiskObject NoteProperty "Server Name" $Server
				Add-Member -InputObject $DiskObject NoteProperty "Volume Name" $disk.name
				Add-Member -InputObject $DiskObject NoteProperty "Total Capacity" $Capacity
				Add-Member -InputObject $DiskObject NoteProperty "Freespace" $Freespace
				Add-Member -InputObject $DiskObject NoteProperty "Percent Available" $PercentAvailable
				
				$DiskCheckResult = $DiskCheckResult += $DiskObject
				
				if ( $PercentAvailable -lt $WarningLevel )
				{
					$DiskCheckError = $DiskCheckError += $DiskObject
				}
			}
		}
		}
        $DiskCheckResult = $DiskCheckResult|sort-object -Property "Percent Available" -Descending
	##End of Check Disk free space...


		## Check CAS Service Status...
		Write-Host "Checking CAS Server Service Status..."  `r
foreach($Server in $CASServers)
{
	$NonStartServices = ""
	$StartServices = ""
	$Services = Test-ServiceHealth $Server.name
	Foreach( $Service in $Services)
	{
		if ($Service.RequiredServicesRunning -ne "true")
		{
			$NonStartServices += $Service.ServicesNotRunning	
			$NonStartServices += " "
		} 
	}
	
	Foreach ( $Service in $CommonServices)
	{
		$s = Get-Service -Name $Service -ComputerName $Server
		if ($s.Status -ne "Running")
		{
			$NonStartServices += $s.DisplayName
			$NonStartServices += " "
		} 
	}
	$CheckTime1 = Get-Date -Format "yyyy-MM-dd HH:mm"
	$ServiceObject = new-Object object
	Add-Member -InputObject $ServiceObject NoteProperty "Check Date" $CheckTime1
	Add-Member -InputObject $ServiceObject NoteProperty "Server Name" $Server
	Add-Member -InputObject $ServiceObject NoteProperty "Not Running Services" $NonStartServices
	$ServiceCheckResult = $ServiceCheckResult += $ServiceObject
	
	if ( $NonStartServices.EndsWith(" ") )
	{
		$ServiceCheckError = $ServiceCheckError += $ServiceObject
	} 
}
		## End of the CAS server Service Check.
		
		
		## Check MBX Service Status...
		Write-Host "Checking MBX Server Service Status..."  `r
	foreach($Server in $MBXServers)
{
	$NonStartServices = ""
	$StartServices = ""
	$Services = Test-ServiceHealth $Server.name
	Foreach( $Service in $Services)
	{
		if ($Service.RequiredServicesRunning -ne "true")
		{
			$NonStartServices += $Service.ServicesNotRunning	
			$NonStartServices += " "
		} 
	}
	
	Foreach ( $Service in $CommonServices)
	{
		$s = Get-Service -Name $Service -ComputerName $Server
		if ($s.Status -ne "Running")
		{
			$NonStartServices += $s.DisplayName
			$NonStartServices += " "
		} 
	}
	
	$s = Get-Service -Name "msftesql-Exchange" -ComputerName $Server.name
		if ($s.Status -ne "Running")
		{
			$NonStartServices += $s.DisplayName
			$NonStartServices += " "
		} 		
	$CheckTime1 = Get-Date -Format "yyyy-MM-dd HH:mm"	
	$ServiceObject = new-Object object
	Add-Member -InputObject $ServiceObject NoteProperty "Check Date" $CheckTime1
	Add-Member -InputObject $ServiceObject NoteProperty "Server Name" $Server.name
	Add-Member -InputObject $ServiceObject NoteProperty "Not Running Services" $NonStartServices
	$ServiceCheckResult = $ServiceCheckResult += $ServiceObject
	
	if ( $NonStartServices.EndsWith(" ") )
	{
		$ServiceCheckError = $ServiceCheckError += $ServiceObject
	} 

}
		## End of MBX Service Check.
	
	## Start to Check Mailbox Database Copy status

	Write-Host "Checking Mailbox Database Copy status..."  `r
	foreach($Server in $MBXServers)
	{
	$DBCopies = Get-MailboxDatabaseCopyStatus -Server $Server
	$CheckTime1 = Get-Date -Format "yyyy-MM-dd HH:mm"
	Foreach($DBCopy in $DBCopies)
	{
		$DBCopyObject = New-Object object
		Add-Member -InputObject $DBCopyObject NoteProperty "Check Date" $CheckTime1
		Add-Member -InputObject $DBCopyObject NoteProperty "Server Name" $Server
		Add-Member -InputObject $DBCopyObject NoteProperty "Database Copy Name" $DBCopy.DatabaseName
		Add-Member -InputObject $DBCopyObject NoteProperty "Status" $DBCopy.Status
                Add-Member -InputObject $DBCopyObject NoteProperty "ContentIndexState" $DBCopy.ContentIndexState
		$DBCopyCheckResult = $DBCopyCheckResult += $DBCopyObject

		if(($DBCopy.Status -ne "Mounted") -and ($DBCopy.Status -ne "Healthy"))
		{
			$DBCopyCheckError = $DBCopyCheckError += $DBCopyObject
		} 
	}
	}
	$DBCopyCheckResult = $DBCopyCheckResult|Sort-Object -Property Status -Descending
	## End of DB Copy Status Check
	
	## Get Database Backup Information ...
$now = [DateTime]::Now
$DBs = Get-MailboxDatabase -Status 

foreach ($DB in $DBs)
{
	Write-Host -ForegroundColor Gray "Checking " $db.name" Backup Information..."
	$LastBackup = @()
	$Ago = @()
	
	if ( $DB.LastFullBackup -eq $null -and $DB.LastIncrementalBackup -eq $null)
	{
		$LastBackupTime = "Never"
		$LastBackupType = "Never"
		[String] $Ago = "Never"
} elseif (($DB.LastFullBackup -eq $null) -or ($DB.LastFullBackup -lt $DB.LastIncrementalBackup))
	{
	$LastBackupTime = $DB.LastIncrementalBackup
		$LastBackupType = "Incremental"
		[int] $Ago = ( $now - $LastBackupTime).TotalHours
		$Ago = "{0:N0}" -f $Ago
	} elseif (($DB.LastIncrementalBackup -eq $null) -or($DB.LastIncrementalBackup -lt $DB.LastFullBackup))
	{
		$LastBackupTime = $DB.LastFullBackup
		$LastBackupType = "Full"
		[int] $Ago = ( $now - $LastBackupTime).TotalHours
		$Ago = "{0:N0}" -f $Ago
	}
	
	$DBObject = New-Object Object
	Add-Member -InputObject $DBObject NoteProperty "Database Name"	$DB.Name
	Add-Member -InputObject $DBObject NoteProperty "Server Name" $DB.Server
	Add-Member -InputObject $DBObject NoteProperty "Backup Type" $LastBackupType
	Add-Member -InputObject $DBObject NoteProperty "Backup Time" $LastBackupTime
	Add-Member -InputObject $DBObject NoteProperty "Elasped time" $Ago
	Add-Member -InputObject $DBObject NoteProperty "Last Full Backup" $DB.LastFullBackup
	Add-Member -InputObject $DBObject NoteProperty "Last Incremental Backup" $DB.LastincrementalBackup
	
	$DBBackupCheckResult = $DBBackupCheckResult += $DBObject
	
	if ($Ago -gt 144 -or $Ago -eq "Never")
	{
		$DBBackupCheckError = $DBBackupCheckError += $DBObject
	}
}
$DBBackupCheckResult = $DBBackupCheckResult|Sort-Object -Property "Elasped time" -Descending
	## End of Database Backup Information Check

## Check Transport server queue status
write-host Checking Queue status...
$Queues = $MBXServers| Get-Queue -Filter {MessageCount -gt 99}
ForEach( $Queue in $Queues)
{
	$QueueObject = New-Object Object
	Add-Member -InputObject $QueueObject NoteProperty "Queue Name" $Queue.Identity
	Add-Member -InputObject $QueueObject NoteProperty "Queue Status" $Queue.Status
	Add-Member -InputObject $QueueObject NoteProperty "Queue Length" $Queue.MessageCount
	Add-Member -InputObject $QueueObject NoteProperty "Last Error" $Queue.LastError
	Add-Member -InputObject $QueueObject NoteProperty "Last Retry Time" $Queue.LastRetryTime
	Add-Member -InputObject $QueueObject NoteProperty "Next Retry Time" $Queue.NextRetryTime 
	$QueueCheckResult = $QueueCheckResult += $QueueObject
	
	if( $Queue.MessageCount -gt 1) # -or $Queue.LastError -ne $Null)
	{
		$QueueCheckError = $QueueCheckError += $QueueObject
	}
}
##End of Checking Transport server queue status

## Check Virtual Directory Service
write-host checking Virtual Directory Service status...
ForEach( $CASServer in $CASServers)
{
	$isSuccess = $NULL
	$EWSTestResults = Test-WebServicesConnectivity -AllowUnsecureAccess -ClientAccessServer $CASServer
	
	foreach( $EWSTest in $EWSTestResults)
	{
		$isSuccess = $EWSTest.Result.Value
		
		$EWSObject = New-Object PSObject
		Add-Member -InputObject $EWSObject NoteProperty "Service" "EWS"
		Add-Member -InputObject $EWSObject NoteProperty "CAS Server" $EWSTest.ClientAccessServerShortName
		Add-Member -InputObject $EWSObject NoteProperty "Server Site" $EWSTest.LocalSiteShortName
		Add-Member -InputObject $EWSObject NoteProperty "Result" $isSuccess
		Add-Member -InputObject $EWSObject NoteProperty "Error" $EWSTest.Error
		Add-Member -InputObject $EWSObject NoteProperty "Latency" $EWSTest.LatencyInMillisecondsString
		Add-Member -InputObject $EWSObject NoteProperty "Scenario" $EWSTest.Scenario
		Add-Member -InputObject $EWSObject NoteProperty "Description" $EWSTest.ScenarioDescription
		
		$EWSTestResult = $EWSTestResult += $EWSObject
		if($isSuccess -ne "Success")
		{
			$EWSTEestError = $EWSTestError += $EWSObject
		}
	}
	
	$IntOWATest = Test-OwaConnectivity -AllowUnsecureAccess -ClientAccessServer $CASServer -TestType Internal
	$ExtOWATest = Test-OwaConnectivity -AllowUnsecureAccess -ClientAccessServer $CASServer -TestType External
	
	$OWAObject = New-Object PSObject
	Add-Member -InputObject $OWAObject NoteProperty "Service" "OWA"
	Add-Member -InputObject $OWAObject NoteProperty "CAS Server" $IntOWATest.ClientAccessServerShortName
	Add-Member -InputObject $OWAObject NoteProperty "Server Site" $IntOWATest.LocalSiteShortName
	Add-Member -InputObject $OWAObject NoteProperty "Result" $IntOWATest.Result.Value
	Add-Member -InputObject $OWAObject NoteProperty "Error" $IntOWATest.Error
	Add-Member -InputObject $OWAObject NoteProperty "Latency" $IntOWATest.LatencyInMillisecondsString
	Add-Member -InputObject $OWAObject NoteProperty "Test Type" $IntOWATest.URLType
	Add-Member -InputObject $OWAObject NoteProperty "Scenario" $IntOWATest.Scenario
	Add-Member -InputObject $OWAObject NoteProperty "Description" $IntOWATest.ScenarioDescription
	$OWATestResult = $OWATestResult += $OWAObject
	if($IntOWATest.Result.Value -ne "Success")
	{
		$OWATestError = $OWATestError += $OWAObject
	}
	
	$OWAObject = New-Object PSObject
	Add-Member -InputObject $OWAObject NoteProperty "Service" "OWA"
	Add-Member -InputObject $OWAObject NoteProperty "CAS Server" $ExtOWATest.ClientAccessServerShortName
	Add-Member -InputObject $OWAObject NoteProperty "Server Site" $ExtOWATest.LocalSiteShortName
	Add-Member -InputObject $OWAObject NoteProperty "Result" $ExtOWATest.Result.Value
	Add-Member -InputObject $OWAObject NoteProperty "Error" $ExtOWATest.Error
	Add-Member -InputObject $OWAObject NoteProperty "Latency" $ExtOWATest.LatencyInMillisecondsString
	Add-Member -InputObject $OWAObject NoteProperty "Test Type" $ExtOWATest.URLType
	Add-Member -InputObject $OWAObject NoteProperty "Scenario" $ExtOWATest.Scenario
	Add-Member -InputObject $OWAObject NoteProperty "Description" $ExtOWATest.ScenarioDescription
	$OWATestResult = $OWATestResult += $OWAObject
	if($ExtOWATest.Result.Value -ne "Success")
	{
		$OWATestError = $OWATestError += $OWAObject
	}
	
	$IntECPTest = Test-EcpConnectivity -ClientAccessServer $CASServer -TestType Internal
	$ExtECPTest = Test-EcpConnectivity -ClientAccessServer $CASServer -TestType External
	$ECPTest = $IntECPTest + $ExtECPTest
	foreach( $T in $ECPTest)
	{
		$ECPObject = New-Object PSObject
		Add-Member -InputObject $ECPObject NoteProperty "Service" "ECP"
		Add-Member -InputObject $ECPObject NoteProperty "CAS Server" $t.ClientAccessServerShortName
		Add-Member -InputObject $ECPObject NoteProperty "Server Site" $t.LocalSiteShortName
		Add-Member -InputObject $ECPObject NoteProperty "Result" $t.Result.Value
		Add-Member -InputObject $ECPObject NoteProperty "Error" $t.Error
		Add-Member -InputObject $ECPObject NoteProperty "Latency" $t.LatencyInMillisecondsString
		Add-Member -InputObject $ECPObject NoteProperty "Test Type" $t.URLType
		Add-Member -InputObject $ECPObject NoteProperty "Scenario" $t.Scenario
		Add-Member -InputObject $ECPObject NoteProperty "Description" $t.ScenarioDescription
		$ECPTestResult = $ECPTestResult += $ECPObject
		if($t.Result.Value -ne "Success")
		{
			$ECPTestError = $ECPTestError += $ECPObject
		}
	}
	
	$ASTest = Test-ActiveSyncConnectivity -AllowUnsecureAccess -ClientAccessServer $CASServer
	foreach( $t in $ASTest)
	{
		$isSuccess = $t.Result.Value
		
		$ASObject = New-Object PSObject
		Add-Member -InputObject $ASObject NoteProperty "Service" "ActiveSync"
		Add-Member -InputObject $ASObject NoteProperty "CAS Server" $t.ClientAccessServerShortName
		Add-Member -InputObject $ASObject NoteProperty "Server Site" $t.LocalSiteShortName
		Add-Member -InputObject $ASObject NoteProperty "Result" $isSuccess
		Add-Member -InputObject $ASObject NoteProperty "Error" $t.Error
		Add-Member -InputObject $ASObject NoteProperty "Latency" $t.LatencyInMillisecondsString
		Add-Member -InputObject $ASObject NoteProperty "Scenario" $t.Scenario
		Add-Member -InputObject $ASObject NoteProperty "Description" $t.ScenarioDescription
		
		$ASTestResult = $ASTestResult += $ASObject
		if($isSuccess -ne "Success")
		{
			$ASTEestError = $ASTestError += $ASObject
		}
	}	
}
## End of Checking Virtual Directory Service





## Check ReplicationHealth
#write-host checking ReplicationHealth....
#$REPLTest = $MBXServers | Test-ReplicationHealth 
#ForEach( $t in $REPLTest)
#{
#	if($t.Server.Contains("ex-cas1")){continue}
#	$REPLObject = New-Object PSObject
#	Add-Member -InputObject $REPLObject NoteProperty "Service" "REPLHealth"
#	Add-Member -InputObject $REPLObject NoteProperty "Server" $t.Server
#	Add-Member -InputObject $REPLObject NoteProperty "Check Result" $t.result
#	Add-Member -InputObject $REPLObject NoteProperty "Check Item" $t.check
#	Add-Member -InputObject $REPLObject NoteProperty "Check Description" $t.CheckDescription
#	Add-Member -InputObject $REPLObject NoteProperty "Error" $t.Error
	
	
#	$REPLTestResult = $REPLTestResult += $REPLObject
#	if( $t.Result.Value -ne "Passed")
#	{
#		$REPLTestError = $REPLTestError += $REPLObject
#	}
#}
## End of checking ReplicationHealth

##Start Count DB users & Database size
Write-Host -ForegroundColor Gray "Checking DB user count..."

foreach ($mdb in $MBXDBs)
{
     $DBpeocountObject = New-Object Object
     $Edbsize=Get-Mailboxdatabase -Status  $mdb|%{$_.databasesize}
	Add-Member -InputObject $DBpeocountObject NoteProperty "Database Name"	$mdb.Name
	Add-Member -InputObject $DBpeocountObject NoteProperty "Database Count" (Get-Mailbox -Database $mdb -ResultSize unlimited).count
        Add-Member -InputObject $DBpeocountObject NoteProperty "Database Size Substring"   $Edbsize.Substring(0,$edbsize.Indexof("(")-1)
	Add-Member -InputObject $DBpeocountObject NoteProperty "Database Size"   $Edbsize.togb()
	$DBPeocountCheckResult = $DBPeocountCheckResult += $DBpeocountObject
}

## End of Counting DB users  Database size

##Start Server ComponentState
Write-Host -ForegroundColor Gray "Checking ServerComponentState..."

$ServerComponentStateResult = get-exchangeserver |get-ServerComponentState |select ServerFqdn, Component, State


## End of Counting ServerComponentState

#HTML styles for nice formatting
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #CC0000; padding: 5px; color: #FFFFFF;}"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
	
#SMTP options for sending the report email
$smtpServer = "10.237.8.100"
$smtpFrom = "systemadmin@xiaomi.com"
$smtpTo = @()
#Get-Content $WorkFolder\$ReportToList
$smtpto +="it-monitor@xiaomi.com"
$smtpto +="systemadmin@xiaomi.com" 
$messageSubject = "RAW Data of XiaoMi Daily Check OPs - "
$messageSubject = $messageSubject += $CheckTime1
$DiskIntro = "<BR><BR>The Disk Free Space Checking Results. <BR><BR>"
$DiskReport = $DiskCheckResult |Sort-Object -Property 'Percent Available' | ConvertTo-Html -Fragment 
$ServiceIntro = "<BR><BR>The Services Checking Results as Below<BR><BR>"
$ServiceReport = $ServiceCheckResult | ConvertTo-Html -Fragment
$DBCopyIntro = "<BR><BR>The DBCopy Status Checking Results as Below<BR><BR>"
$DBCopyReport = $DBCopyCheckResult|Sort-Object -Property 'Status' | ConvertTo-Html -Fragment |foreach {if($_ -like "*<td>Failed</td>*"){$_ -replace "<td>Failed</td>", "<td bgcolor= 'red'>Failed</td>"} else {$_}}
$DBBackupIntro = "<BR><BR>The DB Backup Status is shown below<BR><BR>"
$DBBackupReport = $DBBackupCheckResult | ConvertTo-Html -Fragment
$QueueIntro = "<BR><BR><H1>The Queues Status are listed below<BR><BR>"
$QueueReport = $QueueCheckResult | ConvertTo-Html -Fragment
$TimeDiffIntro = "<BR><BR><H2>The Time Differentiation between Exchange Server and DC.<BR><BR>"
$TimeDiffReport = $TimeDiffCheckResult | ConvertTo-Html -Fragment
$EWSTestIntro = "<BR><BR><H2>EWS Service Test Result<BR><BR>"
$EWSTestReport = $EWSTestResult | ConvertTo-Html -Fragment
$OWATestIntro = "<BR><BR><H2>OWA Service Test Result<BR><BR>"
$OWATestReport = $OWATestResult | ConvertTo-Html -Fragment
$ECPTestIntro = "<BR><BR><H2>ECP Service Test Result<BR><BR>"
$ECPTestReport = $ECPTestResult | ConvertTo-Html -Fragment
$ASTestIntro = "<BR><BR><H2>Active Sync Service Test Result<BR><BR>"
$ASTestReport = $ASTestResult | ConvertTo-Html -Fragment
#$MAPITestIntro = "<BR><BR><H2>MAPI Service Test Result<BR><BR>"
#$MAPITestReport = $MAPITestResult | ConvertTo-Html -Fragment
#$REPLTestIntro = "<BR><BR><H2>DB Copies Replication Health Test Result<BR><BR>"
#$REPLTestReport = $REPLTestResult|Sort-Object -Property "Check Result" | ConvertTo-Html -Fragment |foreach {if($_ -like "*<td>未通过</td>*"){$_ -replace "<td>未通过</td>", "<td bgcolor= 'red'>未通过</td>"} else {$_}}
$DBPeocountIntro = "<BR><BR>The DB People counts is shown below<BR><BR>"
$DBPeocountReport = $DBPeocountCheckResult| Sort-Object -Property "Database Count" -Descending  | ConvertTo-Html -Fragment
$ServerComponentStateIntro = "<BR><BR>The Server ComponState is shown below<BR><BR>"
$ServerComponentStateReport = $ServerComponentStateResult|Sort-Object -Property 'state' -Descending | ConvertTo-Html -Fragment |foreach {if($_ -like "*<td>Inactive</td>*"){$_ -replace "<td>Inactive</td>", "<td bgcolor = 'red'>Inactive</td>"} else {$_}}



#Get ready to send email message
$message = New-Object System.Net.Mail.MailMessage
$message.From = new-object system.net.Mail.Mailaddress $smtpFrom 
if ($smtpTo) 
{
	$smtpTo | foreach {
		$To = new-object system.net.Mail.Mailaddress $_; $Message.To.Add($To)
	}
}
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = ConvertTo-Html -Body "$DBBackupIntro $DBBackupReport $DBCopyIntro $DBCopyReport $DiskIntro $DiskReport $Serviceintro $ServiceReport $QueueIntro $QueueReport $TimeDiffIntro $TimeDiffReport  $DBPeocountIntro $DBPeocountReport $ServerComponentStateIntro $ServerComponentStateReport" -Head $style
$message.body
#Send email message
$smtp = New-Object Net.Mail.SmtpClient -argumentList $smtpServer ;

$smtp.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials ;
write-host sending email...

$smtp.Send($message) ;
write-host sent out email success.
