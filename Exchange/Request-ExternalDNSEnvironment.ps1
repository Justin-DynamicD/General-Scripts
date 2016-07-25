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

#Validate an input has been defined and import
If (!$DomainList -and !$DomainName) {Write-Error "Neither DomainList nor DomainName have been specified.  Nothing to process." -ErrorAction "Stop"}
If ($DomainList -and $DomainName) {Write-Error "You cannot use both DomainList and DomainName simultaneously.  Stoping script." -ErrorAction "Stop"}
If ($DomainList) {$WorkingList = import-csv $CSVFile}
Else {$WorkingList = $DomainName}

#Wipe and set ReturnSet
[System.Collections.ArrayList]$ReturnSet = @()

#Process WorkingList
ForEach ($Name in $WorkingList) {

    #Lookup autodiscover record
    try {
        $baseDiscover = Resolve-DnsName -Name "autodiscover.$Name" -ErrorAction "Stop"
        If ($baseDiscover -eq $null) {$autoDiscover = "Not Found"}
        Else {
            If ([string]$baseDiscover.type[0] -eq "A") {$autoDiscover = "A " + [string]$baseDiscover.IP4Address}
            elseif ([string]$baseDiscover.type[0] -eq "CNAME") {$autoDiscover = "CNAME "+[string]$baseDiscover.NameHost + " (" + [string]$baseDiscover.IP4Address + ")"}
            else {$autoDiscover = "Not Found"}
        } 
    }
    catch {
        $autoDiscover = "Not Found"
    }
    If ($autoDiscover -eq "Not Found") {
        try {
            [string]$baseDiscover = (Resolve-DnsName -Name "_autodiscover._tcp.$Name" -Type SRV -ErrorAction "Stop").NameTarget
            $autoDiscover = "SRV " + $baseDiscover + " (" + (Resolve-DnsName -Name $baseDiscover -ErrorAction "Stop").IP4Address + ")"
        }
        catch {
            $autoDiscover = "Not Found"
        }
    }

    #Lookup MSOID record
    try {
        $msoid = (Resolve-DnsName -Name "msoid.$Name" -Type CNAME -ErrorAction "Stop").NameHost
        If ($msoid -eq $null) {$msoid = "Not Found"} 
    }
        catch {
        $msoid = "Not Found"
    }

    #Lookup spf record
    try {
        $spfRecord = (Resolve-DnsName -Name $Name -Type TXT -ErrorAction "Stop" | Where-Object {$_.Strings -like "v=spf1*"}).Strings
        If ($spfRecord -eq $null) {$spfRecord = "Not Found"}    
    }
    catch {
        $spfRecord = "Not Found"
    }

    #Return Values
    $return = New-Object -TypeName PSCustomObject -Property @{
        DomainName = $Name
        AutoDiscover = $autoDiscover
        MSOID = $msoid
        SPFRecord = $spfRecord
    }

    #Add to ReturnSet ArrayList
    $ReturnSet.add($return)

} #End ForEachLoop

return $ReturnSet