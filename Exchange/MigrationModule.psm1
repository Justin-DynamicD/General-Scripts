
function Initialize-O365User
{
    Param
    (
        # UserName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$true)] 
        [String]$UserName,

         # UserList this is a csv file containing usernames and domains for bulk migrations
        [Parameter(Mandatory=$false)]
        [string]$UserList,

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials

    )

    #Variables specific to client
    $groupList = "Office 365 Enrollment", "Office 365 Enterprise Cal"
    $groupDomain = "corp.ad.viacom.com"
    $searchDomains = "paramount.ad.viacom.com","mtvn.ad.viacom.com","corp.ad.viacom.com"
    $onlineSMTP = "viacom.mail.onmicrosoft.com"

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
                #Write-Verbose $importResults
                [bool]$mSOLActive = $false
                }
            catch {
                write-error "can't switch context to local session" -ErrorAction "Stop"
                }
            }

        #Get CurrentUser and needed SMTP values
        Try {
            $currentUserCount = @()
            ForEach ($domain in $searchDomains) {$currentUserCount += get-aduser -server $domain -filter {name -eq $target} -ErrorAction "Stop"}
            $currentUser = $currentUserCount[0]
            $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
            $newProxy = $currentMailbox.primarysmtpaddress.local + "@"+$onlineSMTP
            }
        Catch {
            write-error "Cannot find either the user account or mailbox for $UserName"
            continue
            }
        
        #Set changes to false
        [bool]$isUpdated = $false

        #Check UPN to primary address
        IF ($currentUser.UserPrincipalName -ne [string]$currentMailbox.primarysmtpaddress) {
            Write-Warning "the UPN and primary SMTP do not match.  Please correct"
        }
        
        #Check for Proxy Address, add if missing
        IF ($currentMailbox.emailAddresses -notcontains $newProxy) {
            Try {
                Write-Verbose "Adding address $newProxy"
                set-mailbox $currentMailbox -Emailaddresses @{add = $newProxy}
                $isUpdated = $true
                }
            Catch {
                write-error "unable to add proxy address, user cannot be moved online!" -ErrorAction "Stop"
                }
        }

        #Check each user to be a member of the groups
        $members=@()
        ForEach ($group in $groupList) {
            try {
                $members = Get-ADGroupMember -Identity $group -server $groupDomain -Recursive | Select -ExpandProperty Name
                }
            Catch {Write-Error "cannot find group $group" -ErrorAction "Stop"}

            If ($members -notcontains $currentUser.Name) {
                Write-Verbose "adding $currentUser.Name to $group"
                Add-ADGroupMember -Identity $group -server $groupDomain -Members $currentUser.Name
                $isUpdated = $true
                } #End Match
            } #End ForEach

        #If any changes were made, output warning
        If ($isUpdated) {
            Write-Warning "groups and/or addresses have been updated for $target, please allow replciation to complete before perfoming actual migration"
            }

    } #End ForEach
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
                [bool]$mSOLActive = $false
                }
            catch {
                write-error "can't switch context to local session" -ErrorAction "Stop"
                }
            }

        #Get CurrentUser and needed SMTP values
        Try {
            $currentUserCount = @()
            ForEach ($domain in $searchDomains) {$currentUserCount += get-aduser -server $domain -filter {name -eq $target} -ErrorAction "Stop"}
            $currentUser = $currentUserCount[0]
            $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
            [string]$primarySMTP = $currentMailbox.primarysmtpaddress
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
                [bool]$mSOLActive = $true
                }
            catch {
                write-error "can't switch context to MSOL session" -ErrorAction "Stop"
                }
            }

        #Do the move
        Write-Verbose "Moving $primarySMTP to O365"
        Try {
            New-MoveRequest -Identity $primarySMTP -Remote -RemoteHostName $RemoteHostName -RemoteCredential $LocalCredentials -TargetDeliveryDomain $targetDeliveryDomain -BadItemLimit 100 -AcceptLargeDataLoss
            While ((Get-MoveRequest -Identity $primarySMTP).Status -eq "InProgress" -or (Get-MoveRequest -Identity $primarySMTP).Status -eq "Queued") {Start-Sleep 15}
            If ((Get-MoveRequest -Identity $primarySMTP).Status -eq "Failed" ) {throw "migration failed, see move request for more details"}
            If ((Get-MoveRequest -Identity $primarySMTP).Status -eq "Suspended" ) {throw "the job has been suspended, see move request for more details"}
            } 
        Catch {
            Write-Error $_.Exception.Message  -ErrorAction "Stop"
            }

        #Update RetentionPolicy
        Write-Verbose "Applying RetentionPolicy to $primarySMTP"
        Set-mailbox -Identity $primarySMTP -RetentionPolicy $retentionPolicy

        #Disable Clutter
        Write-Verbose "Disabling Clutter for $primarySMTP"
        Set-Clutter -Identity $primarySMTP -Enable $false

    } #End per-user Loop
} #End Function