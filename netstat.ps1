#the Visio module is required
if (Get-Module -ListAvailable -Name Visio) {
    Write-Host "Visio module installed - script will proceed."
} else {
    Write-Host "Visio module not installed."
    Write-Host "Will now install Visio module."
    Install-Module Visio
}

Import-Module Visio

#get GeoIP information for ipv4 address using curl
function Get-GeoIP{
    param(
        [string]$ip
    )

    ([xml](curl "http://freegeoip.net/xml/$ip").Content).Response
}

#ask user if geoIP information should be included
$response = Read-Host "Include geoIP information for destination connections? (yes/no)" 
$validResponse = 0 
while($validResponse -eq 0){
    if($response -like 'yes') {
        $getGeoIp = $true
        Write-Host "Will include geoIP information."
        Write-Host "Please be patient for script to gather information."
        $validResponse = 1        
}
            elseif($response -like 'no'){
            Write-Host "GeoIP information will not be included."            
            $getGeoIp = $false
            $validResponse = 1        
            }

            else{
            Write-Host "Please specify 'yes' or 'no'."
            $response = Read-Host "Include geoIP information for destination connections? (yes/no)"
            }
}

#get netstat command output information
function Get-NetStat{
    #get list of local IPs so they can be excluded from Get-GeoIP function
    $localIPs = (gwmi Win32_NetworkAdapterConfiguration | ? { $_.IPAddress -ne $null }).ipaddress
    $localIPs += '127.0.0.1'
   
    #use netstat command to generate custom objects
    Write-Host "Getting netstat information for established connections."
    $cmd = (netstat -ano | Select-String "Established")
    foreach ($line in $cmd){
        Write-Progress -Activity "Working..." -PercentComplete ((($cmd.IndexOf($line) + 1) / $cmd.Count) * 100)
        $cmdline = $line.Line.ToString()
        $elements=$cmdline.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        $object = New-Object -TypeName PSObject
        $object | Add-Member -Name 'Protocol' -MemberType Noteproperty -Value $elements[0]
        $sourceIP = $elements[1].Substring(0, $elements[1].LastIndexOf(':'))
        [array]$global:tempNodes += $sourceIP
        $object | Add-Member -Name 'SourceIP' -MemberType Noteproperty -Value $sourceIP
        $sourcePort = ($elements[1] -split ':')[-1]
        $object | Add-Member -Name 'SourcePort' -MemberType Noteproperty -Value $sourcePort
        $destinationIP = $elements[2].Substring(0, $elements[2].LastIndexOf(':'))
        $global:tempNodes += $destinationIP
        $object | Add-Member -Name 'DestinationIP' -MemberType Noteproperty -Value $destinationIP
        $destinationPort = ($elements[2] -split ':')[-1]
        $object | Add-Member -Name 'DestinationPort' -MemberType Noteproperty -Value $destinationPort
        $object | Add-Member -Name 'State' -MemberType NoteProperty -Value $elements[3]
        $object | Add-Member -Name 'PID' -MemberType NoteProperty -Value $elements[4]
        $object | Add-Member -Name 'ProcessName' -MemberType NoteProperty -Value (ps -Id $object.PID).Name
        [array]$Global:tempEdges += "$sourceIP->$destinationIP"
    
        #probably could do some optimization here
        #only look up if script doesn't already know
        #reference some sort of multi-dim array prior to curl
        if($getGeoIp){
            if($object.DestinationIP -match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' -and $object.DestinationIP -notmatch $localIPs){
                $geoIpInfo = Get-GeoIP($object.DestinationIP)
                $object | Add-Member -Name 'DestLat' -MemberType NoteProperty -Value $geoIpInfo.Latitude
                $object | Add-Member -Name 'DestLong' -MemberType NoteProperty -Value $geoIpInfo.Longitude
                $object | Add-Member -Name 'CountryName' -MemberType NoteProperty -Value $geoIpInfo.CountryName
                $object | Add-Member -Name 'RegionName' -MemberType NoteProperty -Value $geoIpInfo.RegionName
                $object | Add-Member -Name 'City' -MemberType NoteProperty -Value $geoIpInfo.City
            }      
                else{
                    $object | Add-Member -Name 'DestLat' -MemberType NoteProperty -Value 'N/A'
                    $object | Add-Member -Name 'DestLong' -MemberType NoteProperty -Value 'N/A'
                    $object | Add-Member -Name 'CountryName' -MemberType NoteProperty -Value 'N/A'
                    $object | Add-Member -Name 'RegionName' -MemberType NoteProperty -Value 'N/A'
                    $object | Add-Member -Name 'City' -MemberType NoteProperty -Value 'N/A'
                }
        }
        [array]$Global:netstat += $object
     }
}

Get-NetStat

Write-Host "Building node list."
$nodes = $global:tempNodes | sort -Unique
$localIPs = '127.0.0.1|\[::1\]'
$nodes = $nodes | ? {$_ -notmatch $localIPs}
Write-Host "Found" $nodes.Count "nodes."

Write-Host "Building edges list."
$edges = $global:tempEdges | sort -Unique
$localConnections = '127.0.0.1->127.0.0.1|\[::1\]->\[::1\]'
$edges = $edges | ? {$_ -notmatch $localConnections}
Write-Host "Fround" $edges.Count "edges."

Write-Host "Creating Visio objects."
$graph = New-VisioModelDirectedGraph

#create the Visio shapes
$stencil = "basic_u.vssx"
$master = "Circle"
foreach($node in $nodes){
$graph.AddShape($node, $node, $stencil, $master) | Out-Null
}

#create the Visio connectors
foreach($edge in $edges){
    $object = New-Object -TypeName PSObject
    $object | Add-Member -Name 'ID' -MemberType Noteproperty -Value $edge
    $object | Add-Member -Name 'from' -MemberType Noteproperty -Value ($edge.split("->")[0])
    $object | Add-Member -Name 'to' -MemberType Noteproperty -Value ($edge.split("->")[-1])
    $object | Add-Member -Name 'label' -MemberType Noteproperty -Value $edge
    $object | Add-Member -Name 'type' -MemberType NoteProperty -Value curved
    $object | Add-Member -Name 'begin_arrow' -MemberType NoteProperty -Value 0
    $object | Add-Member -Name 'end_arrow' -MemberType NoteProperty -Value 4
    $object | Add-Member -Name 'hyperlink' -MemberType NoteProperty -Value $null

    [array]$connections += $object
}

foreach($connection in $connections){
$graph.AddConnection($connection.ID, $connection.from, $connection.to, $connection.label, $connection.type, $connection.begin_arrow, $connection.end_arrow, $connection.hyperlink) | Out-Null
}

New-VisioApplication

Write-Host "Printing Visio graph."
Out-Visio $graph