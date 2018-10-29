#Requires -Version 3.0
<#

    .SYNOPSIS
        report a subset of all VMs properties on hostIP host(s) and export them to html and/or Excel.

    .DESCRIPTION
        Retrieves a list of virtual machines from hostIP hosts or vCenter Server, extracting a subset of 
        vm properties values returned by Get-View. Hosts names(IPs) and credentials are saved in a XMl file ( config.xml for this script). 
        The script saves reports to disk and automatically displays the reports converted to HTML by invoking the default browser.
        Reports are exported to Excel using Import-Excel Modules from https://github.com/dfinke/ImportExcel.
        You can install this Module from PowerShell Gallery : https://www.powershellgallery.com/packages/ImportExcel/5.3.4

    .USAGE 

        .\listvms <sortBy>:[Optional] <exportTo>:[optional] 
    
    .PARAMETER sortBY
        sort result : [optional]
        sort result by -ramalloc -ramhost -os 
    .PARAMETER exportTo
        export result : [optional]
        export result to -excel -html (default is both)

    .EXAMPLE
        .\listAllVMs 
        .\listAllVMs -exportTo excel -os  
    .NOTES
        Author: Arnaud Mutana
        Last Updated: OCTOBER 2018
        Version: 2.1
  
    .Requires :
    VMware Infrastructure
    ImportExcel Modules
#> 
# commandline parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false, Position = 1)]
    [string]$sortBy,
    [Parameter(Mandatory = $false, Position = 2)]
    [string]$exportTo
)

# Load required Snapins and Modules
if ($null -eq (Get-PSSnapin -Name VeeamPSSNapin -ErrorAction SilentlyContinue)) {
    Add-PSSnapin VeeamPSSNapin
}
if ($null -eq (Get-Module -Name Import -ErrorAction SilentlyContinue)) {
    Import-Module "$PSScriptRoot\ImportExcel"
}


#Populate PSObject with the required vm properties 
function vmProperties {
    param([PSObject]$view)
 
    $list = foreach ($vm in $view) {
        #State info
        if ($vm.Runtime.PowerState -eq "poweredOn") {$state = "ON"}
        elseif ($vm.Runtime.PowerState -eq "poweredOff") {$state = "OFF"}
        else {$state = "n/a"}
        #VMtools state
        if ($vm.summary.guest.ToolsRunningStatus -eq "guestToolsRunning") {$vmtools = "Running"}
        elseif ($vm.summary.guest.ToolsRunningStatus -eq "guestToolsNotRunning") {$vmtools = "Not running"}
        else {$vmtools = "n/a"}
    
        # Net Info
        $ipAdresses = $vm.Guest.IpAddress
        $macAdresses = $vm.Guest.Net.MacAddress
        #Datastores info
        $datastoresNames = $vm.Config.DataStoreUrl.Name
        #Check for multi-homed vms
        $ips=""
        foreach ($ip in $ipAdresses){
            $ips += $ip + "`n"
        }
        $macs=""
        foreach($mac in $macAdresses){
            $macs += $mac + "`n"
        }
        $datastores=""
        foreach($datastore in $datastoresNames){
            $datastores += $datastore + " "
        }

        #Disk Info 
        $disks = $vm.guest.disk
        #sum up disk space
        $diskName = ""
        $capacity = 0
        $freeSpace = 0
        foreach ($disk in $disks) {
            $diskName += "[" + $disk.diskpath + "], " 
            $capacity += [Decimal]([math]::round($disk.capacity / 1GB, 2))
            $freeSpace += [decimal]([math]::round($disk.freespace / 1GB, 2))
        }

        if ($diskName.Length -gt 0) {
            $diskName = $diskName.Substring(0, $diskName.Length - 2)
        }

        #Populate object
        [PSCustomObject]@{
            "Name"      = $vm.Name
            "OS"        = $vm.Guest.GuestFullName
            "Hostname"  = $vm.summary.guest.hostname
            "vCPUs"     = $vm.Config.hardware.NumCPU
            "Cores"     = $vm.Config.Hardware.NumCoresPerSocket
            "RAM Alloc" = $vm.Config.Hardware.MemoryMB
            "RAM Host"  = $vm.summary.QuickStats.HostMemoryUsage
            "RAM guest" = $vm.summary.QuickStats.GuestMemoryUsage
            "Disk"      = $diskName
            "Capacity"  = $capacity
            "Free"      = $freeSpace
            "NICS"      = $vm.Summary.config.NumEthernetCards
            "IPs"       = $ips
            "MACs"      = $macs 
            "Datastore" = $datastores 
            "vmTools"   = $vmtools
            "State"     = $state
            "UUID"      = $vm.Summary.config.Uuid
            "VM ID"     = $vm.Summary.vm.value
        }
    }
 
    return $list
}
 
#Stylesheet - this is used by the ConvertTo-html cmdlet
function header {
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

# XML file parse
[xml]$XmlDoc = Get-Content "$PSScriptRoot\Config.xml"
$listOfESX = $XmlDoc.Config.ESX


#############################
### Script entry point ###
#############################   
foreach ($ESX in $listOfESX) {
    $hostIP = $ESX.HostIP.ip
    $user = $ESX.User.user
    $pass = $ESX.password.pass

    #Path to html report
    $repPath = (Get-ChildItem  env:userprofile).value + "\desktop\{0}.htm" -f $hostIP
    $excelPath = (Get-ChildItem  env:userprofile).value + "\desktop\{0}.xlsx" -f $hostIP
    
    #Report Title
    $title = "<h2>VMs hosted on $hostIP</h2>"
    
    #Sort by
    if ($sortBy -eq "") {$sortBy = "Name"; $desc = $False} 
    elseif ($sortBy.Equals("ramalloc")) {$sortBy = "RAM Alloc"; $desc = $True} 
    elseif ($sortBy.Equals("ramhost")) {$sortBy = "RAM Host"; $desc = $True} 
    elseif ($sortBy.Equals("os")) {$sortBy = "OS"; $desc = $False}
    
    Try {
        #Drop any previously established connections
        Disconnect-VIServer -Confirm:$False -ErrorAction SilentlyContinue
    
        #Connect to vCenter or hostIP
        if (($user -eq "") -or ($pass -eq "")) 
        {Connect-VIServer $hostIP -ErrorAction Stop}
        else 
        {Connect-VIServer $hostIP -User $user -Password $pass -ErrorAction Stop}
    
        #Get a VirtualMachine view of all vms
        $vmView = Get-View -viewtype VirtualMachine

        #export to
        if ($exportTo -eq "excel") {
            #Iterate through the view object, write the set of vm properties to a PSObject and convert the whole lot to Excel workbook
            (vmProperties -view $vmView) | Sort-Object -Property @{Expression = $sortBy; Descending = $desc} | Export-Excel -Path $excelPath
        }
        elseif ($exportTo -eq "html") {
            #Iterate through the view object, write the set of vm properties to a PSObject and convert the whole lot to Excel workbook
            (vmProperties -view $vmView) | Sort-Object -Property @{Expression = $sortBy; Descending = $desc} | ConvertTo-Html -Head $(header) -PreContent $title | Set-Content -Path $repPath -ErrorAction Stop
        }
        else {
            (vmProperties -view $vmView) | Sort-Object -Property @{Expression = $sortBy; Descending = $desc} | ConvertTo-Html -Head $(header) -PreContent $title | Set-Content -Path $repPath -ErrorAction Stop
            (vmProperties -view $vmView) | Sort-Object -Property @{Expression = $sortBy; Descending = $desc} | ConvertTo-Html -Head $(header) -PreContent $title | Set-Content -Path $repPath -ErrorAction Stop
        }
        #Disconnect from vCenter or hostIP
        Disconnect-VIServer -Confirm:$False -Server $hostIP -ErrorAction Stop
        #Load report in default browser
        Invoke-Expression "cmd.exe /C start $repPath"

    }
    Catch {
        Write-Host $_.Exception.Message
    }
}