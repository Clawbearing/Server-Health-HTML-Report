$ostitle = ConvertTo-Html -Fragment -PreContent '<h2>Windows Version<h2/>'
$osversion = [Environment]::OSVersion | Select-Object -Property Version
$osversion.Version
$osfrag = $osversion 

$hotfixtitle2 = ConvertTo-Html -Fragment -PreContent '<h2>Last Updated Installed on<h2/>'
$hotfix = Get-HotFix | Sort-Object -Descending -Property InstalledOn | Select-Object -First 1
$($hotfix.InstalledOn) 
$hotfixfrag = $hotfix | ConvertTo-Html -Fragment

$uptimetitle = ConvertTo-Html -Fragment -PreContent '<h2>Computer Uptime<h2/>'
$uptime = (get-date) - (gcim Win32_OperatingSystem).LastBootUpTime | Select-Object -Property Days, Hours, Minutes
$uptime2 = "Computer has been up for the past $($uptime.Days) days, $($uptime.Hours) hours, and $($uptime.Minutes) minutes"
$uptimefrag = $uptime2 | ConvertTo-Html -Fragment


$workingsetprocess = Get-Process | Select-Object -Property WorkingSet, ProcessName, Id -Last 10 | Sort-Object -Property WorkingSet -Descending 
$cpuprocess = Get-Process | Select-Object -Property CPU, ProcessName, Id -Last 10 | Sort-Object -Property CPU -Descending



#load percentage of CPU
$cputotalloadaverage = Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select Average

$diskqueueslength = Get-Counter -Counter '\PhysicalDisk(*)\Avg. Disk Queue Length'


$diskspace = Get-CimInstance -Class CIM_LogicalDisk | Select-Object @{Name="Size(GB)";Expression={[math]::Round( $_.size/1gb)}}, @{Name="Free Space(GB)";Expression={[math]::Round( $_.freespace/1gb)}}, @{Name="Free (%)";Expression={"{0,6:P0}" -f(($_.freespace/1gb) / ($_.size/1gb))}}, DeviceID, DriveType | Where-Object DriveType -EQ '3'


#Total Memory
$totalmemory = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb


#memory available
$memoryavail = Get-Counter -Counter '\Memory\Available Bytes'
$memoryround = [math]::floor($memoryavail.CounterSamples.cookedvalue /1GB)


$systemlogs = Get-EventLog -LogName System -EntryType Error -Newest 10
$applogs = Get-EventLog -LogName Application -EntryType Error -Newest 10



#logs
$syslogsfrag = $systemlogs | ConvertTo-HTML -Property TimeGenerated, InstanceID, Message -Fragment

$applogfrag = $applogs | ConvertTo-HTML -Property TimeGenerated, InstanceID, Message -Fragment



 #processes doing the most stuff
 $wksfrag = $workingsetprocess | ConvertTo-HTML -Property WorkingSet, ProcessName, ID -Fragment
 $cpufrag = $cpuprocess | ConvertTo-HTML -Property CPU, ProcessName, ID -Fragment

 #load percentage of CPU
 $cpuloadtotal = "$($cputotalloadaverage.Average)% total CPU usage"

 #free space on drives
 $dkspacefrag = $diskspace | ConvertTo-Html -Fragment

 #Total Memory
 $totmem = "$totalmemory GB of total memory"

 #memory available
 $availemem = "$memoryround GB of memory available"

 #network table
$networkthroughput = Get-WmiObject -class Win32_PerfFormattedData_Tcpip_NetworkInterface|select Packetspersec, CurrentBandwidth, BytesReceivedPersec, BytesSentPersec | where {$_.PacketsPersec -gt 0}

$netfrag = $networkthroughput | ConvertTo-Html -Fragment

$space = @() | ConvertTo-Html -Fragment -PreContent ‘<h2> </h2>’ 
$space2 = @() | ConvertTo-Html -Fragment -PreContent ‘<h2> </h2>’ 

$reporttitle = ConvertTo-Html -Fragment -PreContent '<h1>$ComputerName $DATE </h1>'

$disktitle = ConvertTo-Html -Fragment -PreContent '<h2>Disk Space Storage<h2/>'

$generalstats = ConvertTo-Html -Fragment -PreContent '<h2>Memory and CPU Usage<h2/>'

$systemlogstitle = ConvertTo-Html -Fragment -PreContent '<h2>System ERROR Logs<h2/>'

$applogstitle = ConvertTo-Html -Fragment -PreContent '<h2>Application ERROR Logs<h2/>'

$processesmemorytitle = ConvertTo-Html -Fragment -PreContent '<h2>Most Used Processes (Memory)<h2/>'

$processescputitle = ConvertTo-Html -Fragment -PreContent '<h2>Most Used Processes (CPU)<h2/>'

$netutiliztitle = ConvertTo-Html -Fragment -PreContent '<h2>Network Utilization (May Be Empty/Blank)<h2/>'

$ReprotTest = ConvertTo-Html -Body "$reporttitle $space $space2 $ostitle $space2 $($osversion.Version) $space2 $hotfixtitle $space2 $($hotfix.InstalledOn) $space2 $uptimetitle $space2 $uptime2 $space2 $disktitle $space $space2 $dkspacefrag $space $space2 $generalstats $space $space2 $totmem $space $space2 $availemem $space $space2 $cpuloadtotal $space $space2 $systemlogstitle $space $space2 $syslogsfrag $space $space2 $applogstitle $space $space2 $applogfrag $space $space2 $processesmemorytitle $space $space2 $wksfrag $space $space2 $processescputitle $space $space2 $cpufrag $space $space2 $netutiliztitle $space $space2 $netfrag" -Title COMPUTER

$ReprotTest | Out-File -FilePath .\Documents\lastfrag.html

Invoke-Item -Path .\Documents\lastfrag.html


