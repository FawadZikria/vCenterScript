$date = Get-Date -Format F

if ( $DefaultVIServers.Length -lt 1 )
{
  Connect-VIServer -Server $server -Protocol https -User $user -Password $pass -WarningAction SilentlyContinue | Out-Null
}

# Format html report
$htmlReport = @"
<style type='text/css'>
.heading {
 	color:#3366FF;
	font-size:12.0pt;
	font-weight:700;
	font-family:Verdana, sans-serif;
	text-align:left;
	vertical-align:middle;
	height:30.0pt;
	width:416pt
}
.colnames {
 	color:white;
	font-size:8.0pt;
	font-weight:700;
	font-family:Tahoma, sans-serif;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid windowtext;
	background:#153E7E;
}
.text {
	color:#FFFFFF;
	font-size:10.0pt;
	font-family:Arial;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid windowtext;
	background:#7F525D;
}

</style>
<table border=0 cellpadding=2 cellspacing=2 width=1000
 style='border-collapse:collapse;table-layout:auto;width:650pt'>
 <tr style='height:17.0pt'>
  <th colspan=6 class="text">Cluster's Capacity DashBoard - Deployment </th>
  <th colspan=2 class="text">$date</th>
 <tr>
 <tr>
	<th class="text">ClusterName</th>
	<th class="text">Total No. of Hosts</th>
   <th class="text">Total CPU Sockets</th>
   <th class="text">Total CPU Cores</th>
    <th class="text">Total Cluster's Physical Memory</th>
   <th class="text">Total Cluster's Physical Memory Available</th>
   <th class="text">MEM Percentage Available ( Considering HA 1 host Failure ) </th>
   <th class="text">Total Virtual Machines</th>
   </tr>
   
"@

#-----------------Getting RDM & Storage Details -------------------------------------

$totalRDMSize = 0
$RDvms = Get-VM  | Get-View
foreach($vm in $RDvms){
  foreach($dev in $vm.Config.Hardware.Device){
    if(($dev.gettype()).Name -eq "VirtualDisk"){
       if(($dev.Backing.CompatibilityMode -eq "physicalMode") -or 
          ($dev.Backing.CompatibilityMode -eq "virtualMode")){
           $totalRDMSize = $totalRDMSize + $dev.CapacityInKB
           }
     }
  }
}

#$StorageTotal = (Get-Datastore |select CapacityMB | Measure-Object CapacityMB -Sum).sum
$StorageTotal = (Get-Datastore | where {$_.Extensiondata.Summary.MultipleHostAccess} | select CapacityMB | Measure-object CapacityMB -Sum).sum
#$StorageAvail = (Get-Datastore |select freeSpaceMB | Measure-Object freeSpaceMB -Sum).sum
$StorageAvail = (Get-Datastore | where {$_.Extensiondata.Summary.MultipleHostAccess} | select FreeSpaceMB | Measure-object FreeSpaceMB -Sum).sum

#$StorageTotal = [math]::round($StorageTotal * 1MB / 1TB)
$StorageTotal = [math]::round(($StorageTotal /1024/1024) ,2)
#$StorageAvail = [math]::round($StorageAvail * 1MB / 1TB)
$StorageAvail = [math]::round(($StorageAvail /1024/1024) ,2)
#$totalRDMSize = [math]::round($totalRDMSize * 1KB / 1TB)
$totalRDMSize = [math]::round(($totalRDMSize /1024/1024/1024) ,2)

#-----------------Getting RDM & Storage Details -------------------------------------

$htmlReport1 = @"
<style type='text/css'>
.text {
	color:#FFFFFF;
	font-size:10.0pt;
	font-family:Arial;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid windowtext;
	background:#4E8975;
}

</style>
<table border=0 cellpadding=3 cellspacing=3 width=2000
 style='border-collapse:collapse;table-layout:auto;width:650pt'>
 <tr></tr>
 <tr style='height:17.0pt'>
	  <th colspan=7 class="text">Storage Utilization</th>
  </tr>
    
   <tr style='height:17.0pt'>
	  <th colspan=2 class="text">Total Storage ( Excluding RDMs )</th>
	   <th colspan=1 class="text">$StorageTotal TB</th>
	    <th colspan=1 class="text">Total RDM Size </th>
	   <th colspan=1 class="text">$totalRDMSize TB</th>
	   <th colspan=1 class="text">Total Free Storage Available</th>
	   <th colspan=1 class="text">$StorageAvail TB</th>
  </tr>
  <tr></tr>

 	<tr style='height:17.0pt'>
	  <th colspan=7 class="text">Cluster's DashBoard - Resource Utilization</th>
	   
  </tr>
	<th class="text">ClusterName</th>
	<th class="text">Total Effective CPU Available</th>
   <th class="text">CPU Usage MAX (last 7 Days)</th>
   <th class="text">CPU Usage Average (last 7 Days)</th>
   <th class="text">Total Effective Memory</th>
   <th class="text">MEM Usage MAX (last 7 Days)</th>
   <th class="text">MEM Usage Average (last 7 Days)</th>
   </tr>
   

   
"@

$htmlReport2 = @"
<style type='text/css'>
.text {
	color:#FFFFFF;
	font-size:10.0pt;
	font-family:Arial;
	text-align:center;
	vertical-align:middle;
	border:.5pt solid windowtext;
	background:#4E8975;
}

</style>
<table border=0 cellpadding=3 cellspacing=3 width=2000
 style='border-collapse:collapse;table-layout:auto;width:650pt'>
 <tr></tr>
 
  <tr style='height:17.0pt'>
	  <th colspan=8 class="text">Cluster's DashBoard - Resource Allocation</th>
	   
  </tr>
	<th class="text">ClusterName</th>
	<th class="text">Total Poweron VMs </th>
    <th class="text">Total Virtual Cores </th>
   <th class="text">Total Virtual Cores Allocated</th>
   <th class="text">total Physical Memory</th>
   <th class="text">Total Virtual Memory Allocated</th>
   <th class="text">Total disks Allocated</th>
   <th class="text">Total disks Utilization</th>
   </tr>
     
"@

#-----------------Getting Physical CPU & Memory for each cluster  -------------------------------------

$cluster = Get-Cluster | sort name

foreach ($clustername in $cluster)
{
$hostlist = Get-cluster $Clustername | get-vmhost | sort name	
$hostcount = $hostlist.Count
$totalVMs = (Get-Cluster $clustername | Get-VM).count
$totalClusMEM = 0
$totalClusCPU = 0
$totalClusCore = 0
$memoryEffectiveUsed = 0
$cpuEffectiveUsed = 0
foreach($name in $hostlist)
{
$totalClusMEM = $totalClusMEM + (($name | Get-view ).Hardware.MemorySize)
$totalClusCPU = $totalClusCPU + (($name | get-view ).Hardware.CpuInfo.NumCpuPackages)
$totalClusCore = $totalClusCore + (($name | Get-View).Hardware.CpuInfo.NumCpuCores)
}

#$statcpuMAX = Get-Stat -Entity ($clustername)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat cpu.usagemhz.maximum
$statcpu = Get-Stat -Entity ($clustername)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat cpu.usagemhz.average -IntervalSecs 1800
#$statmemMAX = Get-Stat -Entity ($clustername)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat mem.usage.maximum
$statmem = Get-Stat -Entity ($clustername)-start (get-date).AddDays(-7) -Finish (Get-Date)-MaxSamples 10000 -stat mem.usage.average -IntervalSecs 1800

#$cpuMAX = $statcpuMAX | Measure-Object -Property value -Maximum -Average
$cpu = $statcpu | Measure-Object -Property value -Average -Maximum
#$memMAX = $statmemMAX | Measure-Object -Property value -Maximum
$mem= $statmem | Measure-Object -Property value -Average -Maximum

$cpuMAX = [Math]::Round(($cpu.Maximum / 1000),2)
$cpuAVG = [Math]::Round(($cpu.Average / 1000),2)
$memMAX = [Math]::Round($mem.Maximum,2)
$memAVG = [Math]::Round($mem.Average,2)

#$memmax = [math]::round(($mem.Maximum /1024/1024) , 2)
#$memaverage = [math]::round(($mem.Average /1024/1024) ,2)

$vmlist = get-vm -location $clustername | select name, NumCPU , MemoryMB, PowerState, ProvisionedSpaceGB, UsedSpaceGB
$totalCPU = 0
$totalMEM = 0
$PoweronVM = 0
$PoweroffVM = 0
$TotalStorageProvision = 0
$totalstorageUsed = 0

ForEach ($vm in $vmlist)
{
$totalCPU = $totalCPU + $vm.numcpu
$totalMEM = $totalMEM + $vm.MemoryMB
$TotalStorageProvision = $TotalStorageProvision + $vm.ProvisionedSpaceGB
$totalstorageUsed = $totalstorageUsed + $vm.UsedSpaceGB


if ($vm.PowerState -eq 'PoweredOn')
{
$PoweronVM = $PoweronVM + 1
}
else
{
$PoweroffVM = $PoweroffVM + 1
}
}

$totalMEM = [math]::round($totalMEM * 1MB / 1GB)
$totalClusMEM = [math]::round($totalClusMEM / 1GB, 0)
#$totalCoreAvail = $totalClusCore - $totalCPU
$totalMEMAvail = $totalClusMEM - $totalMEM

$1hostfailMEM = ($totalClusMEM - ( $totalClusMEM / $hostcount ))
$1hostfailMEMAvail = ($totalMEMAvail - ( $totalClusMEM / $hostcount ))
$totalMEMprcen = [Math]::Round((($1hostfailMEMAvail / $1hostfailMEM)*100),2)

#$1hostfailCPU = ($totalClusCore - ( $totalClusCore / $hostcount ))
#$1hostfailCores = ($totalCoreAvail - ( $totalClusCore / $hostcount ))
#$totalCPUPrcen = [Math]::Round((($1hostfailCores / $1hostfailCPU)*100),2)

$ClusterResource = get-cluster $clustername | Get-View
$ClusterMEMmb = [Math]::Round(($ClusterResource.Summary.EffectiveMemory / 1024),2)
$ClusterCPUmz = [Math]::Round(($ClusterResource.Summary.EffectiveCpu / 1000),2)

$htmlReport = $htmlReport +
			  "<td class='text'>" + $Clustername + "</td>" +
			  "<td class='text'>" + $hostcount + "</td>" + 
			  "<td class='text'>" + $totalClusCPU + "</td>" +
			  "<td class='text'>" + $totalClusCore + "</td>" +
			  "<td class='text'>" + $totalClusMEM + ' GB' + "</td>"+
			  "<td class='text'>" + $totalMEMAvail + ' GB' + "</td>"+
			  "<td class='text'>" + $totalMEMprcen + ' %' + "</td>"+
			   "<td class='text'>" + $totalVMs + "</td></tr>"
			   
$htmlReport1 = $htmlReport1 +
			  "<td class='text'>" + $Clustername + "</td>" +
			  "<td class='text'>" + $ClusterCPUmz + ' GHz' +  "</td>" +
			  "<td class='text'>" + $cpumax + ' GHz' + "</td>" +
			  "<td class='text'>" + $cpuAVG + ' GHz' + "</td>" +
			  "<td class='text'>" + $ClusterMEMmb + ' GB' + "</td>" +
			  "<td class='text'>" + $memMAX + ' %' + "</td>" +
			  "<td class='text'>" + $memAVG + ' %' + "</td></tr>"
			
$htmlReport2 = $htmlReport2 +
			  "<td class='text'>" + $Clustername + "</td>" +
			  "<td class='text'>" + $PoweronVM +  "</td>" +
			  "<td class='text'>" + ($totalClusCore *4) + "</td>" +
			  "<td class='text'>" + $totalCPU + "</td>" +
			  "<td class='text'>" + $totalClusMEM + ' GB' + "</td>" +
			  "<td class='text'>" + $totalMEM + ' GB' + "</td>" +
              "<td class='text'>" + [math]::round($TotalStorageProvision , 2 ) + ' GB' + "</td>" +
			  "<td class='text'>" + [math]::round($totalstorageUsed , 2 )  + ' GB' + "</td></tr>"	  
}



$htmlReport = $htmlReport + "</table>"
$htmlReport1 = $htmlReport1 + "</table>"
$htmlReport2 = $htmlReport2 + "<table>"
$htmlReport3 = $htmlReport + $htmlReport1 + $htmlReport2
