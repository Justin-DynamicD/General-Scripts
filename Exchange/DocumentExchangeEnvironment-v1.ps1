#Function to add Resolve-DNSName to Win7 Machines
function Resolve-DnsName2008
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [string]$Server = '127.0.0.1'
    )
    Try
    {
        $nslookup = &nslookup.exe $Name $Server
        $regexipv4 = "^(?:(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)\.){3}(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)$"

        $name = @($nslookup | Where-Object { ( $_ -match "^(?:Name:*)") }).replace('Name:','').trim()

        $deladdresstext = $nslookup -replace "^(?:^Address:|^Addresses:)",""
        $Addresses = $deladdresstext.trim() | Where-Object { ( $_ -match "$regexipv4" ) }

        $total = $Addresses.count
        $AddressList = @()
        for($i=1;$i -lt $total;$i++)
        {
            $AddressList += $Addresses[$i].trim()
        }

        $AddressList | %{

        new-object -typename psobject -Property @{
            Name = $name
            IPAddress = $_
            }
        }
    }
    catch 
    { }
}

#Check for ResolveDNSName and create the alias if it is missing
if (!(Get-Command Resolve-DNSName -ErrorAction SilentlyContinue)) {
    new-alias -name Resolve-DNSName -value Resolve-DnsName2008 -description "Resolve-DNSName"
    [switch]$removealias = $True  
}

#region Load Exchange Snap-in
if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue))
{
	if (Test-Path "C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1")
	{
		. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
		Connect-ExchangeServer -auto
	} elseif (Test-Path "C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1") {
		Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.Admin
		.'C:\Program Files\Microsoft\Exchange Server\bin\Exchange.ps1'
	} else {
		throw "Exchange Management Shell cannot be loaded"
	}
}
#endregion Load Exchange Snap-in

#region HTML Start
$HTML='<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>IAMGOLD Exchange Environment</title>
</head>

<body>
<h2 align="center" a id=Top">Exchange Environment</h2>
<h4 align="center">Generated ' + $((Get-Date).ToString()) + '</h4>
<hr>
<h4>Table of contents</4><br>
<a href="#ExchangeServersInformation">Exchange Server Information</a><br>
<a href="#DAGInformation">Database Availability Group Information</a><br>
<a href="#DatabaseInformation">Databases Information</a><br>
<a href="#CASInformation">Client Access Server Information</a><br>
<a href="#CASArrayInformation">CAS Array Information</a><br>
<a href="#TransportServersInformation">Transport Server Information</a><br>
<a href="#TransportCosts">Transport Costs</a><br>
<a href="#SendConnectors">Send Connectors</a><br>
<a href="#AcceptedDomains">Accepted Domains</a><br>
<a href="#EmailPolicy">Email Address Policies</a><br>
'


#endregion HTML Start

#region Get Exchange Servers
$servers = Get-ExchangeServer | Sort-Object Site

$HTML +='<h3><a id="ExchangeServersInformation"><b>Exchange Servers Information</b></a></h3>
<table border="0" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Server Name</h4></td>
		<td style="height: 23px"><h4>Version</h4></td>
        <td style="height: 23px"><h4>Site</h4></td>
		<td style="height: 23px"><h4>IP Address</h4></td>
		<td style="height: 23px"><h4>Roles</h4></td>
	</tr>
'

foreach ($exchangeserver in $servers)
{
 $ServerName = $exchangeserver.name.ToUpper()
 $ServerMajorVersion = $exchangeserver.AdminDisplayVersion.Major
 if ($ServerMajorVersion -eq 14)
 {
  $ServerVersion = "Exchange 2010"
 }
 else
 {
  $ServerVersion = "Exchange 2013"
 }
$ServerVersion += " (SP" + $exchangeServer.AdminDisplayVersion.Minor.ToString() + ") Build " + $exchangeServer.AdminDisplayVersion.Build.ToString() + "." + $exchangeServer.AdminDisplayVersion.Revision.ToString()
$ServerSite = $exchangeserver.site.Name.ToString()
$IP = Resolve-DnsName $exchangeserver.Name.ToString()
$ServerIP = $IP.IPAddress.ToString()
$ServerRoles = $exchangeserver.ServerRole.ToString()

$HTML += '
    <tr>
		<td style="height: 23px">' + $ServerName + '</td>
		<td style="height: 23px">' + $ServerVersion + '</td>
		<td style="height: 23px">' + $ServerSite + '</td>
		<td style="height: 23px">' + $ServerIP + '</td>
        <td style="height: 23px">' + $ServerRoles + '</td>
    </tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Get Exchange Servers

#region Get DAG Information
$DAGs = Get-DatabaseAvailabilityGroup | Sort-Object Name

$HTML +='<h3><a id="DAGInformation"><b>Database Availability Groups Information</b></a></h3>
<table border="1" cellpadding="5" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>DAG Name</h4></td>
		<td style="height: 23px"><h4>IP Addresses</h4></td>
        <td style="height: 23px"><h4>Members</h4></td>
		<td style="height: 23px"><h4>Witness Server</h4></td>
		<td style="height: 23px"><h4>Witness Directory</h4></td>
	</tr>
'
foreach ($DAG in $DAGs)
{
$DAGName = $DAG.Name
$DAGIP = $DAG.DatabaseAvailabilityGroupIpv4Addresses
$DAGMembers = $DAG.Servers
$DAGWitnessServer = $DAG.WitnessServer
$DAGWitnessDirectory = $DAG


$HTML += '
    <tr>
		<td style="height: 23px">' + $DAGName + '</td>
		<td style="height: 23px">' + $DAGIP + '</td>
		<td style="height: 23px">' + $DAGMembers + '</td>
		<td style="height: 23px">' + $DAGWitnessServer + '</td>
        <td style="height: 23px">' + $DAGWitnessDirectory + '</td>
    </tr>
    <tr>
        <td align="center" colspan="5">
           Mailbox Databases
           <table border="2">
            <tr>
                <td><h4>Database Name</h4></td>
                <td><h4>RPC CAS Server</h4></td>
                <td><h4>Hosting Server</h4></td>
                <td><h4>EDB File Path</h4></td>
                <td><h4>Log Folder Path</h4></td>
                <td><h4>Circular Logging</h4></td>
                <td><h4>Copies (Activation Preference)</h4></td>
            </tr>'

$DAGDatabases = Get-MailboxDatabase | where {$_.MasterServerOrAvailabilityGroup -eq $dag.name}
foreach ($DAGDB in $DAGDatabases)
{
$DBName = $DAGDB.Name
$DBRPCCAS = $DAGDB.RPCClientAccessServer
$DBHostingServer = $DAGDB.Server.Name
$DBEDB = $DAGDB.EDBFilePath.PathName
$DBLog = $DAGDB.LogFolderPath.PathName
$DBCircularLogging = $DAGDB.CircularLoggingEnabled
$HTML +='<tr>
            <td>' + $DBName + '</td>
            <td>' + $DBRPCCAS + '</td>
            <td>' + $DBHostingServer + '</td>
            <td>' + $DBEDB + '</td>
            <td>' + $DBLog + '</td>
            <td>' + $DBCircularLogging + '</td>
            <td><table>'

$DBCopies = $DAGDB.ActivationPreference
foreach ($DBCopy in $DBCopies)
{
    $DBCopyHost = $DBCopy.Key.Name
    $DBCopyPriority = $DBCopy.Value
    $HTML += '<tr><td>' + $DBCopyHost + ' (' + $DBCopyPriority + ')</td></tr>'
}
$HTML += '</table></td></tr>'
}

$HTML +='           </table>
        </td>
    </tr>

'
}
$HTML +='</table>
<br><a href="#Top">Back to top</a>
<hr>'

#endregion Get Dag Information

#region Get Database Information
$Databases = Get-MailboxDatabase | Sort-Object ServerName
$HTML +='<h3><a id="DatabaseInformation"><b>Databases Information</b></a></h3>
<table border="1" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Database Name</h4></td>
		<td style="height: 23px"><h4>Hosted On</h4></td>
        <td style="height: 23px"><h4>Issue Warning Quota</h4></td>
		<td style="height: 23px"><h4>Prohibit Send Quota</h4></td>
		<td style="height: 23px"><h4>Prohibit Send/Receive Quota</h4></td>
        <td style="height: 23px"><h4>Deleted Item Retention</h4></td>
        <td style="height: 23px"><h4>Offline Address Book</h4></td>
        <td style="height: 23px"><h4>Public Folder Database</h4></td>
	</tr>
'

foreach ($Database in $Databases)
{
$DBName = $Database.Name
$DBServer = $Database.ServerName
$DBWarn = $Database.IssueWarningQuota
$DBSQ = $Database.ProhibitSendQuota
$DBSRQ = $Database.ProhibitSendReceiveQuota
$DBDeletedItems = $Database.DeletedItemRetention.Days
$DBOAB = $Database.OfflineAddressBook
$DBPF = $Database.PublicFolderDatabase

$HTML += '
    <tr>
		<td style="height: 23px">' + $DBName + '</td>
		<td style="height: 23px">' + $DBServer + '</td>
		<td style="height: 23px">' + $DBWarn + '</td>
		<td style="height: 23px">' + $DBSQ + '</td>
        <td style="height: 23px">' + $DBSRQ + '</td>
		<td style="height: 23px">' + $DBDeletedItems + '</td>
        <td style="height: 23px">' + $DBOAB + '</td>
        <td style="height: 23px">' + $DBPF + '</td>
    </tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Get Database Information

#region Get CAS information
$CASServers = Get-ClientAccessServer

$HTML +='<h3><a id=CASInformation"><b>Client Access Servers Information</b></a></h3>
<table border="1" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Server Name</h4></td>
		<td style="height: 23px"><h4>AutoDiscover Site</h4></td>
        <td style="height: 23px"><h4>Outlook Anywhere Enabled</h4></td>
	</tr>
'

foreach ($CASServer in $CASServers)
{
$CASServerName = $CASServer.Name
$CASSite = $CASServer.AutoDiscoverSiteScope
$CASOAEnabled = $CASServer.OutlookAnywhereEnabled
$HTML +='
    <tr>
		<td style="height: 23px">' + $CASServerName + '</td>
		<td style="height: 23px">' + $CASSite + '</td>
		<td style="height: 23px">' + $CASOAEnabled + '</td>
    </tr>
    <tr>
    <td align="center" colspan="3">OWA Virtual Directory
    <table>
        <tr>
            <td>Name</td>
            <td>Internal Authentication Method</td>
            <td>Internal URL</td>
            <td>External URL</td>
        </tr>'
$OWAVirtualDirectories = Get-OwaVirtualDirectory -Server $CASServerName
foreach ($OWADirectory in $OWAVirtualDirectories)
{
$OWAVDName = $OWADirectory.name
$OWAVDAuth = $OWADirectory.InternalAuthenticationMethods
$OWAVDIntURL = $OWADirectory.InternalUrl
$OWAVDExtURL = $OWADirectory.ExternalUrl

$HTML += '
    <tr>
		<td style="height: 23px">' + $OWAVDName + '</td>
		<td style="height: 23px">' + $OWAVDAuth + '</td>
		<td style="height: 23px">' + $OWAVDIntURL + '</td>
        <td style="height: 23px">' + $OWAVDExtURL + '</td>
    </tr>
'
}
    
$HTML +='</table></td></tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Get CAS Information

#region Get CAS Array
$CASArrays = Get-ClientAccessArray | Sort-Object SiteName

$HTML +='<h3><a id="CASArrayInformation"><b>CAS Array Information</b></a></h3>
<table border="1" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Name</h4></td>
		<td style="height: 23px"><h4>FQDN</h4></td>
        <td style="height: 23px"><h4>IP</h4></td>
        <td style="height: 23px"><h4>Site</h4></td>
		<td style="height: 23px"><h4>Members</h4></td>
		<td style="height: 23px"><h4>Used by Databases</h4></td>
	</tr>
'

foreach ($CASArray in $CASArrays)
{
$CASArrayName = $CASArray.Name
$CASArrayFQDN = $CASArray.Fqdn
$CASArrayIP = Resolve-DnsName $CASArrayFQDN.ToString()
$CASArraySite = $CASArray.SiteName
$CASArrayMembers = $CASArray.Members
$CASArrayDBs = Get-MailboxDatabase | where {$_.RpcClientAccessServer -eq $CASArrayFQDN}
 
$HTML += '
    <tr>
		<td style="height: 23px">' + $CASArrayName + '</td>
		<td style="height: 23px">' + $CASArrayFQDN + '</td>
		<td style="height: 23px">' + $CASArrayIP + '</td>
		<td style="height: 23px">' + $CASArraySite + '</td>
		<td style="height: 23px">' + $CASArrayMembers + '</td>
        <td style="height: 23px">'
$Count = $CASArrayDBs.Count
Do
{
$HTML += $CASArrayDBs[$Count - 1].Name
$Count = $Count - 1
If ($Count -ne 0)
{
$HTML += '<br>'
}
}
Until ($Count -eq 0)
$HTML += '</td>
    </tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Get CAS Array

#region Get Transport Servers
$HubServers = Get-TransportServer | Sort-Object Name

$HTML +='<h3><a id="TransportServersInformation"><b>Transport Servers Information</b></a></h3>
<table border="1" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
'
foreach ($HubServer in $HubServers)
{
$HTML +='
	<tr>
		<td style="height: 23px"><h4>Server Name</h4></td>
        <td style="height: 23px">' + $Hubserver.Name + '</td>
	</tr>
'
$ReceiveConnectors = Get-ReceiveConnector -Server $HubServer.Name
$HTML += '
<tr>
    <td align="center" colspan="2">Receive Connectors
        <table>
            <tr>
                <td><b>Name</b></td>
                <td><b>Authentication Mechanism</b></td>
                <td><b>Bindings</b></td>
                <td><b>Permission Groups</b></td>
                <td><b>Max Message Size</b></td>
                <td><b>Requires TLS</b></td>
            </tr>
'
foreach ($ReceiveConnector in $ReceiveConnectors)
{
$ConnectorName = $ReceiveConnector.Name
$ConnectorAuth = $ReceiveConnector.AuthMechanism
$ConnectorBindings = $ReceiveConnector.Bindings
$ConnectorMax = $ReceiveConnector.MaxMessageSize
$ConnectorPermission = $ReceiveConnector.PermissionGroups
$ConnectorTLS = $ReceiveConnector.RequireTLS

$HTML += '
    <tr>
		<td style="height: 23px">' + $ConnectorName + '</td>
		<td style="height: 23px">' + $ConnectorAuth + '</td>
		<td style="height: 23px">' + $ConnectorBindings + '</td>
        <td style="height: 23px">' + $ConnectorPermission + '</td>
        <td style="height: 23px">' + $ConnectorMax + '</td>
        <td style="height: 23px">' + $ConnectorTLS + '</td>
    </tr>
'
}
$HTML += '
</table></td></tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Get Transport Servers

#region Get Transport Costs
$HTML +='<h3><a id="TransportCosts"><b>Transport Costs</b></a></h3>
<table border="1" cellpadding="5">
<tr>
<td></td>
'
$Servers = Get-ExchangeServer | where {$_.IsHubTransportServer -eq $True} | Sort-Object Site
$ServerCount = 0
Do{
$HTML +='
<td>' + $Servers[$ServerCount].Name + '</td>
'
$ServerCount ++
}
Until ($ServerCount -eq $Servers.Count-1)


$HTML +='
</tr>
'

foreach ($Server in $Servers)
{
$Count = 0
$ServerSite = $Server.Site.Name

$HTML +='
<tr>
<td align="right">' + $Server.Name + '</td>
'
Do{
$RemoteServerSite = $Servers[$Count].Site.Name
If ($Server.Name -eq $Servers[$Count].Name)
{
$HTML +='<td align="center">0</td>'
}
ElseIf ($ServerSite -eq $RemoteServerSite)
{
$HTML +='<td align="center">0</td>'
}
Else
{
$ADSiteLinkCost = ((get-adsitelink) | where {($_.Sites).name -like $ServerSite -and ($_.Sites).name -like $RemoteServerSite}).adcost
If ($ADSiteLinkCost) {
If ($ADSiteLinkCost -eq 10) {$color = "green"}
Else {$color = "yellow"}
$HTML +='<td align="center" style="background-color:' + $color + '">' + $ADSiteLinkCost +'</td>'}
Else {$HTML +='<td style="background-color:black"></td>'}
}
$Count ++
}
Until ($Count -eq $Servers.Count -1)

$HTML +='</tr>'

}
$HTML +='</table>
<br>
<table border="1">
<tr>
<td align="center">0</td>
<td>Local MAPIDevlivery</td>
</tr>
<tr>
<td align="center" style="background-color:green">10</td>
<td>Main Route</td>
</tr>
<tr>
<td align="center" style="background-color:yellow">30-45</td>
<td>Backup Route</td>
</tr>
<tr>
<td align="center" style="background-color:black"></td>
<td>No Direct Route</td>
</tr>
</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Transport Costs

#region Get Send Connector
$SendConnectors = Get-SendConnector

$HTML +='<h3><a id="SendConnectors"><b>Send Connectors</b></a></h3>
<table border="0" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Connector Name</h4></td>
		<td style="height: 23px"><h4>Address Spaces</h4></td>
        <td style="height: 23px"><h4>Source Servers</h4></td>
        <td style="height: 23px"><h4>Smart Hosts</h4></td>
		<td style="height: 23px"><h4>MaxMessageSize</h4></td>
	</tr>
'

foreach ($SendConnector in $SendConnectors)
{
$HTML += '
    <tr>
		<td style="height: 23px">' + $SendConnector.Name + '</td>
		<td style="height: 23px">'
            foreach ($AddressSpace in $SendConnector.AddressSpaces)
            {
                $Count = 1
                $HTML += $AddressSpace.Type + ' : ' + $AddressSpace.Address + ' ; ' + $AddressSpace.Cost
                if ($Count -ne $AddressSpace.count){$HTML += '<br>'}
                $Count++
            }
$HTML +='</td>
		<td style="height: 23px">'
            foreach ($SourceServer in $SendConnector.SourceTransportServers)
            {
                $Count = 1
                $HTML += $SourceServer.Name
                if ($Count -ne $SourceServer.count){$HTML += '<br>'}
                $Count++
            }
$HTML +='</td>
		<td style="height: 23px">' + $SendConnector.SmartHostsString + '</td>
		<td style="height: 23px">' + $SendConnector.MaxMessageSize.Value.ToMB() + ' MB</td>
    </tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Get External Connector

#region Accepted Domain Name
$AcceptedDomains = Get-AcceptedDomain

$HTML +='<h3><a id="AcceptedDomains"><b>Accepted Domains</b></a></h3>
<table border="0" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Domain Name</h4></td>
		<td style="height: 23px"><h4>Domain Type</h4></td>
        <td style="height: 23px"><h4>Default</h4></td>
	</tr>
'

foreach ($Domain in $AcceptedDomains)
{
$HTML += '
    <tr>
		<td style="height: 23px">' + $Domain.DomainName + '</td>
        <td style="height: 23px">' + $Domain.DomainType + '</td>
        <td style="height: 23px">' + $Domain.Default + '</td>
    </tr>
'
}
$HTML += '</table>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Accepted Domain Name

#region Email Address Policies
$EmailPolicies = Get-EmailAddressPolicy

$HTML +='<h3><a id="EmailPolicy"><b>Email Address Policies</b></a></h3>
<table border="0" cellpadding="3" style="font-size:8pt;font-family:Arial,sans-serif">
	<tr>
		<td style="height: 23px"><h4>Policy Name</h4></td>
		<td style="height: 23px"><h4>Recipient Filter</h4></td>
        <td style="height: 23px"><h4>Email Address Templates</h4></td>
	</tr>
'

foreach ($Policy in $EmailPolicies)
{
$HTML += '
    <tr>
		<td style="height: 23px">' + $Policy.Name + '</td>
        <td style="height: 23px">' + $Policy.RecipientFilter + '</td>
        <td style="height: 23px"><table><tr><td>Address Template</td><td>Type</td><td>Primary</td>'
        foreach ($Template in $Policy.EnabledEmailAddressTemplates)
        {
                $HTML += '<tr><td>' + $Template.AddressTemplateString + '</td><td>' + $Template.Prefix + '</td><td>' + $Template.IsPrimaryAddress.ToString() + '</td></tr>'
        }
$HTML +=        '</table></td>
    </tr>
'
}
$HTML += '</table><br>
<table border="1">
    <tr>
        <td>%g</td>
        <td>First Name</td>
    </tr>
    <tr>
        <td>%s</td>
        <td>Last Name</td>
    </tr>
    <tr>
        <td>%m</td>
        <td>Alias</td>
    </tr>
    <tr>
        <td>%1g</td>
        <td>First letter of the First Name</td>
    </tr>
<br><a href="#Top">Back to top</a>
<hr>'
#endregion Email Address Policies

#region End HTML File
$HTML += '
</body>
</html>
'
#endregion End HTML File

#Write HTML file
echo $HTML > test.html

If ($removealias) {
    remove-item alias:Resolve-DNSName
}  
