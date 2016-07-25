#Request-ExternalDNSEnvironment
Param
(
    [Parameter(Mandatory=$False)]
    [string]$DomainName,

    [Parameter(Mandatory=$False)]
    [string]$DomainList,

    [Parameter(Mandatory=$false)]
    [string]$CSVFile = ".\DNSResults.csv"
)

#Validate an input has been defined
If (!$DomainList -and !$DomainName) {Write-Error "Neither DomainList nor DomainName have been specified.  Nothing to process." -ErrorAction "Stop"}
If ($DomainList -and $DomainName) {Write-Error "You cannot use both DomainList and DomainName simultaneously.  Stoping script." -ErrorAction "Stop"}

#Lookup autodiscover record
try {
    $baseDiscover = Resolve-DnsName -Name "autodiscover.$DomainName" -ErrorAction "Stop"
    If ($baseDiscover -eq $null) {$autoDiscover = "Not Found"}
    Else {
        $autoDiscover = [string]$baseDiscover.type[0]+" "+[string]$baseDiscover.IP4Address
    } 
}
catch {
    $autoDiscover = "Not Found"
}
If ($autoDiscover -eq "Not Found") {
    try {
        [string]$baseDiscover = (Resolve-DnsName -Name "_autodiscover._tcp.$DomainName" -Type SRV -ErrorAction "Stop").NameTarget
        $autoDiscover = "SRV "+(Resolve-DnsName -Name $baseDiscover -ErrorAction "Stop").IPAddress
    }
    catch {
        $autoDiscover = "Not Found"
    }
}

#Lookup MSOID record
try {
    $msoid = (Resolve-DnsName -Name "msoid.$DomainName" -Type CNAME -ErrorAction "Stop").NameHost
    If ($msoid -eq $null) {$msoid = "Not Found"} 
}
    catch {
    $msoid = "Not Found"
}

#Lookup spf record
try {
    $spfRecord = (Resolve-DnsName -Name $DomainName -Type TXT -ErrorAction "Stop" | Where-Object {$_.Strings -like "v=spf1*"}).Strings
    If ($spfRecord -eq $null) {$spfRecord = "Not Found"}    
}
catch {
    $spfRecord = "Not Found"
}

#Return Values
$return = New-Object -TypeName PSCustomObject -Property @{
    AutoDiscover = $autoDiscover
    MSOID = $msoid
    SPFRecord = $spfRecord
}
return $return