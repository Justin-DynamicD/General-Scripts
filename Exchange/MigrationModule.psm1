<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Initialize-O365User
{
    Param
    (
        # UserName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$true)] 
        [String]$UserName,

        # DomainName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$false)] 
        [String]$DomainName = "paramount.ad.viacom.com"
    )

    #Variables specific to client
    $groupList = "Office 365 Enrollment", "Office 365 Enterprise Cal"
    $groupDomain = "corp.ad.viacom.com"
    $onlineSMTP = "viacom.mail.onmicrosoft.com"

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure htey are avialable" -ErrorAction "Stop"}
        }

    #Get CurrentUser and needed SMTP values
    Try {
        Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"
        $currentUser = get-aduser -server $DomainName -filter {name -eq $UserName} -ErrorAction "Stop"
        $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
        $newProxy = "smtp:"+$UserName+"@"+$onlineSMTP
        }
    Catch {
        write-error "Cannot find either the user account or mailbox for $UserName" -ErrorAction "Stop"
        }


    #Check UPN to primary address
    IF ($currentUser.UserPrincipalName -ne $currentMailbox.primarysmtpaddress) {
         Write-Warning "the UPN and primary SMTP do not match.  Please correct"
    }
    
    #Check for Proxy Address, add if missing
    
    IF ($currentUser.proxyAddresses -notcontains $newProxy) {
        Write-Verbose "Adding address $newProxy"
        set-ADUser $UserName -add proxyAddresses = $newProxy
    }

    #Check each user to be a member of the groups
    $members=@()
    [bool]$isUpdated = $false
    ForEach ($group in $groupList) {
        try {
            $members = Get-ADGroupMember -Identity $group -server $groupDomain -Recursive | Select -ExpandProperty Name
            }
        Catch {Write-Error "cannot find group $group" -ErrorAction "Stop"}

        If ($members -notcontains $UserName) {
            Write-Verbose "adding $UserName to $group"
            Add-ADGroupMember -Identity $group -server $groupDomain -Members $UserName
            $isUpdated = $true
            } #End Match
        } #End ForEach

    #If any changes were made, output warning
    If ($isUpdated) {
        Write-Warning "groups have been updated, please allow replciation to complete before perfoming actual migration"
        }  
} #End Function


function Move-O365User {

    param (
        # UserName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$true)]
        [string]$UserName,

        # DomainName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$false)] 
        [String]$DomainName = "paramount.ad.viacom.com",

        # RemoteHostName These are valid endpoints for replicating to the cloud
        [Parameter(Mandatory=$false)][ValidateSet("owa.viacom.com","owa.mtvne.com","mail.paramount.com")] 
        [string]$RemoteHostName = "owa.viacom.com",

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [pscredential]$OnlineCredentials,

        # LocalCredentials These are the credentials require to sign into your Exchange Environment
        [Parameter(Mandatory=$true)]
        [pscredential]$LocalCredentials
    )

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure htey are available" -ErrorAction "Stop"}
        }

    #Get CurrentUser and needed SMTP values
    Try {
        Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"
        $currentUser = get-aduser -server $DomainName -filter {name -eq $UserName} -ErrorAction "Stop"
        $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
        $primarySMTP = $currentMailbox.primarysmtpaddress
        }
    Catch {
        write-error "Cannot find either the user account or mailbox for $UserName" -ErrorAction "Stop"
        }

    #Grab current SendLimits
    If ($currentMailbox.UseDatabaseQuotaDefaults) {
        $dB = get-mailbox $currentMailbox.database.name
        IF ($dB.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {$DBReceiveQuota = $dB.ProhibitSendReceiveQuota.Value}
        IF ($dB.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {$DBSendQuota = $dB.ProhibitSendQuota.Value}
        IF ($dB.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {$DBWarning = $dB.IssueWarningQuota.Value}

        Write-Verbose "Updating Storage Quotas"
        set-mailbox $currentUser.Name -ProhibitSendQuota $DBSendQuota -ProhibitSendReceiveQuota $DBReceiveQuota -IssueWarningQuota $DBWarning

        }
    
    <#
    Else {
        IF ($currentMailbox.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {$DBReceiveQuota = $currentMailbox.ProhibitSendReceiveQuota.Value}
        IF ($currentMailbox.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {$DBSendQuota = $currentMailbox.ProhibitSendQuota.Value}
        IF ($currentMailbox.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {$DBWarning = $currentMailbox.IssueWarningQuota.Value}    
        }
    #>

    #Update Send Quotas
    

    #Connect to the Exchange online environment and clobber all modules
    [bool]$mSOLActive = $false
    $search = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($search -ne $NULL) {[bool]$mSOLActive = $true}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            $importResults = Import-PSSession $mSOLSession -AllowClobber
            Write-Verbose $importResults
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    Else {
        $importResults = Import-PSSession $search -AllowClobber
        }
    
    #Variables specific to client
    $targetDeliveryDomain = "viacom.mail.onmicrosoft.com"

    # Do the move
    Write-Verbose "Moving user to O365"
    New-MoveRequest -Identity $primarySMTP -Remote -RemoteHostName $RemoteHostName -RemoteCredential $LocalCredentials -TargetDeliveryDomain $targetDeliveryDomain -BadItemLimit 100 -AcceptLargeDataLoss

    #Disable Clutter
    Write-Verbose "Disabling Clutter"
    Set-Clutter -Identity $primarySMTP -Enable $false
    
} #End Function