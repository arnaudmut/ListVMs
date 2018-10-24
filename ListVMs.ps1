<#

    .SYNOPSIS
        List all VM on hostIP host

    .DESCRIPTION
        Retrieves a list of virtual machines from an hostIP host or vCenter Server, extracting a subset of 
        vm properties values returned by Get-View. The list is converted to HTML and written to disk. 
        The script automatically displays the report by invoking the default browser.

    .USAGE 

        .\listvms <vcenter or host ip>:[Manadatory] <user>:[Manadatory] <password>:[Manadatory] <sortBy>:[Optional]

    .PARAMETER hostIP
        vcenter or host ip:[Manadatory]
    
    .PARAMETER user
        user login: [Manadatory] 
    
    .PARAMETER pass
        password : user password [Manadatory] 
    
    .PARAMETER sortBY
        sort result : [optional]
        sort result by -ramalloc -ramhost -os 

    .EXAMPLE
        .\listVMs 172.17.10.100 root mypassword ramalloc  
    .NOTES
        Author: Arnaud Mutana
        Last Updated: SEPTEMBER 2018
        Version: 1
  
    .Requires :
    Veeam Backup & Replication v9.5 Update 3 (full or console install)
    VMware Infrastructure

#> 
 
#Command line parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,Position=1)]
    [string]$hostIP,
  [Parameter(Mandatory=$false,Position=2)]
  [string]$user,

 [Parameter(Mandatory=$false,Position=3)]
 [string]$pass,

 [Parameter(Mandatory=$false,Position=4)]
 [string]$sortBy
 
)

#Populate PSObject with the required vm properties 
function vmProperties
{
 param([PSObject]$view)
 
 $list=foreach ($vm in $view){
  #State info
  if ($vm.Runtime.PowerState -eq "poweredOn") {$state="ON"}
   elseif ($vm.Runtime.PowerState -eq "poweredOff") {$state="OFF"}
    else {$state="n/a"}
  #VMtools state
  if ($vm.summary.guest.ToolsRunningStatus -eq "guestToolsRunning") {$vmtools="Running"}
   elseif ($vm.summary.guest.ToolsRunningStatus -eq "guestToolsNotRunning") {$vmtools="Not running"}
    else {$vmtools="n/a"}
    
    # Net Info
 $ips = $vm.Guest.IpAddress
 $macs = $vm.Guest.Net.MacAddress
 #Check for multi-homed vms - max. 2 ips
  if ($ips.count -gt 1)
  {$ips=$vm.guest.net.ipaddress[0] + " " + $vm.guest.net.ipaddress[1]} 
 
  if ($macs.count -gt 1)
   {$macs=$vm.guest.net.macaddress[0] + " " + $vm.guest.net.macaddress[1]} 
 
   #Disk Info 
 $disks = $vm.guest.disk
  #Check for multi-homed vms - max. 2 ips

  $diskName = ""
  $capacity = 0
  $freeSpace = 0
   foreach ($disk in $disks)
 {
     $diskName +=  "[" + $disk.diskpath + "], " 
     $capacity +=  [Decimal]([math]::round($disk.capacity/1GB, 2))
     $freeSpace += [decimal]([math]::round($disk.freespace/1GB, 2))
 }

if($diskName.Length -gt 0)
{
    $diskName = $diskName.Substring(0,$diskName.Length-2)
}

 #Populate object
 [PSCustomObject]@{
  "Name" = $vm.Name
  "OS" = $vm.Guest.GuestFullName
  "Hostname" = $vm.summary.guest.hostname
  "vCPUs" = $vm.Config.hardware.NumCPU
  "Cores" = $vm.Config.Hardware.NumCoresPerSocket
  "RAM Alloc" = $vm.Config.Hardware.MemoryMB
  "RAM Host" = $vm.summary.QuickStats.HostMemoryUsage
  "RAM guest" = $vm.summary.QuickStats.GuestMemoryUsage
  "Disk" = $diskName
  "Capacity" = $capacity
  "Free" = $freeSpace
  "NICS" = $vm.Summary.config.NumEthernetCards
  "IPs" = $ips
  "MACs" = $macs  
  "vmTools" = $vmtools
  "State" = $state
  "UUID" = $vm.Summary.config.Uuid
  "VM ID" = $vm.Summary.vm.value
  }
 }
 
 return $list
}
 
#Stylesheet - this is used by the ConvertTo-html cmdlet
function header{
 $style = @"
 <style>
 body{
 font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
 }
 
 table{
  border-collapse: collapse;
  border: none;
  font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
  color: black;
  margin-bottom: 10px;
  table-layout: fixed;
 }
 
 table td{
  font-size: 10px;
  padding-left: 0px;
  padding-right: 20px;
  text-align: left;
  break-word: break-word;
 }
 
 table th{
  font-size: 10px;
  font-weight: bold;
  padding-left: 0px;
  padding-right: 20px;
  text-align: left;
 }
 
 h2{
  clear: both; font-size: 130%;color:#00134d;
 }
 
 p{
  margin-left: 10px; font-size: 12px;
 }
 
 table.list{
  float: left;
 }
 
 table tr:nth-child(even){background: #e6f2ff;} 
 table tr:nth-child(odd) {background: #FFFFFF;}
 
 div.column {width: 320px; float: left;}
 div.first {padding-right: 20px; border-right: 1px grey solid;}
 div.second {margin-left: 30px;}
 
 table{
  margin-left: 10px;
 }
 –>
 </style>
"@
 
 return [string] $style
 }

#############################
### Script entry point ###
#############################
#Path to html report
 $repPath=(Get-ChildItem  env:userprofile).value+"\desktop\{0}.htm" -f $hostIP
 
#Report Title
 $title = "<h2>VMs hosted on $hostIP</h2>"
 
#Sort by
 if ($sortBy -eq "") {$sortBy="Name"; $desc=$False} 
  elseif ($sortBy.Equals("ramalloc")) {$sortBy = "RAM Alloc"; $desc=$True} 
   elseif ($sortBy.Equals("ramhost")) {$sortBy = "RAM Host"; $desc=$True} 
    elseif ($sortBy.Equals("os")) {$sortBy = "OS"; $desc=$False}
 
Try{
 #Drop any previously established connections
  Disconnect-VIServer -Confirm:$False -ErrorAction SilentlyContinue
 
 #Connect to vCenter or hostIP
  if (($user -eq "") -or ($pass -eq "")) 
   {Connect-VIServer $hostIP -ErrorAction Stop}
  else 
    {Connect-VIServer $hostIP -User $user -Password $pass -ErrorAction Stop}
 
 #Get a VirtualMachine view of all vms
  $vmView = Get-View -viewtype VirtualMachine
 
 #Iterate through the view object, write the set of vm properties to a PSObject and convert the whole lot to HTML
  (vmProperties -view $vmView) | Sort-Object -Property @{Expression=$sortBy;Descending=$desc} | ConvertTo-Html -Head $(header) -PreContent $title | Set-Content -Path $repPath -ErrorAction Stop
 
 #Disconnect from vCenter or hostIP
  Disconnect-VIServer -Confirm:$False -Server $hostIP -ErrorAction Stop
   #Load report in default browser
 Invoke-Expression "cmd.exe /C start $repPath"

}
Catch
 {
  Write-Host $_.Exception.Message
 }
