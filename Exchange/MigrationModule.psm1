
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

    param (
        # UserName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$false)]
        [string]$UserName,

        # UserList this is a csv file containing usernames and domains for bulk migrations
        [Parameter(Mandatory=$false)]
        [string]$UserList,

        # RemoteHostName These are valid endpoints for replicating to the cloud
        [Parameter(Mandatory=$false)][ValidateSet("owa.viacom.com","owa.mtvne.com","mail.paramount.com")] 
        [string]$RemoteHostName = "owa.viacom.com",

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials,

        # LocalCredentials These are the credentials require to sign into your Exchange Environment
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$LocalCredentials
    )

    #Variables specific to client
    $targetDeliveryDomain = "viacom.mail.onmicrosoft.com"
    $searchDomains = "paramount.ad.viacom.com","mtvn.ad.viacom.com"

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    
    #Import UserList into a workingList
    If ($UserList) {$workingList = (import-csv $UserList -header UserName).UserName}
    Else {$workingList = $UserName}

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure they are available" -ErrorAction "Stop"}
        }
    Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"

    #Connect to the Exchange online environment and track all cmdlets
    [bool]$mSOLActive = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -ne "ps.outlook.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($mSOLSession -ne $NULL) {[bool]$mSOLActive = $true}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    
    #Begin per-user Loop
    ForEach ($target in $workingList) {

        #Set Current Session to Local Host
        IF ($mSOLActive) {
            Try {
                $importResults = Import-PSSession $localSession -AllowClobber
                Write-Verbose $importResults
                }
            catch {
                write-error "can't switch context to local session" -ErrorAction "Stop"
                }
            }

        #Get CurrentUser and needed SMTP values
        Try {
            $currentUser = @()
            ForEach ($domain in $searchDomains) {$currentUser = get-aduser -server $domain -filter {name -eq $target} -ErrorAction "Stop"}
            If ($currentUser.count -ne 1) {write-Error -Message "Username found " + $currentUser.count + " matches.  Should only be 1" -ErrorAction "Stop"}
            $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
            $primarySMTP = $currentMailbox.primarysmtpaddress
            }
        Catch {
            write-error "Cannot find either the user account or mailbox for $target" -ErrorAction "Stop"
            }

        #Grab current SendLimits and RetentionPolicy
        $retentionPolicy = $currentMailbox.RetentionPolicy.Name
        If ($currentMailbox.UseDatabaseQuotaDefaults) {
            write-verbose "$UserName storage quotas are being pulled form the database, updating storage quotas"
            $dB = get-mailbox $currentMailbox.database.name
            IF ($dB.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {$DBReceiveQuota = $dB.ProhibitSendReceiveQuota.Value}
            IF ($dB.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {$DBSendQuota = $dB.ProhibitSendQuota.Value}
            IF ($dB.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {$DBWarning = $dB.IssueWarningQuota.Value}

            set-mailbox $currentUser.Name -ProhibitSendQuota $DBSendQuota -ProhibitSendReceiveQuota $DBReceiveQuota -IssueWarningQuota $DBWarning
            }

        #switch session to mSOL
        IF (!$mSOLActive) {
            Try {
                $importResults = Import-PSSession $mSOLSession -AllowClobber
                Write-Verbose $importResults
                }
            catch {
                write-error "can't switch context to MSOL session" -ErrorAction "Stop"
                }
            }

        # Do the move
        Write-Verbose "Moving $primarySMTP to O365"
        New-MoveRequest -Identity $primarySMTP -Remote -RemoteHostName $RemoteHostName -RemoteCredential $LocalCredentials -TargetDeliveryDomain $targetDeliveryDomain -BadItemLimit 100 -AcceptLargeDataLoss

        #Update RetentionPolicy
        Write-Verbose "Applying RetentionPolicy to $primarySMTP"
        Set-mailbox -Identity $primarySMTP -RetentionPolicy $retentionPolicy

        #Disable Clutter
        Write-Verbose "Disabling Clutter for $primarySMTP"
        Set-Clutter -Identity $primarySMTP -Enable $false

    } #End per-user Loop
} #End Function