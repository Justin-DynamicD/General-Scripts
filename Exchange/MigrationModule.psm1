
function Initialize-O365User
{
    Param
    (
        # UserName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$false)] 
        [String]$UserName,

         # UserList this is a csv file containing usernames and domains for bulk migrations
        [Parameter(Mandatory=$false)]
        [string]$UserList,

        # SettingsOutFile This specifies the filename to store check-result data into
        [Parameter(Mandatory=$false)] 
        [string]$SettingsOutFile = ".\MailboxSettings "+ (get-date -format m) + ".csv",

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials

    )

    #Variables specific to client
    $groupList = "Office 365 Enrollment", "Office 365 Enterprise Cal" #Ensures users are a membe of listed groups.  THese have been identified as used for assigning licenses
    $groupDomain = "corp.ad.viacom.com" #Domains that above groups are members of
    $globalCatalog = "jumboshrimp.mtvn.ad.viacom.com:3268"
    $onlineSMTP = "viacom.mail.onmicrosoft.com"

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    
    #Import UserList into a workingList
    If ($UserList) {$workingList = Get-Content $UserList}
    Else {$workingList = $UserName}

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure they are available" -ErrorAction "Stop"}
        }
    Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"

    #Connect to the Exchange online environment and track all cmdlets
    [bool]$mSOLActive = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($mSOLSession -ne $NULL) {[bool]$mSOLActive = $true}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    
    # create an Empty settings log before check
    [System.Collections.ArrayList]$settingsOutLog = @()

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
            $currentUser = @()
            $currentUser += get-aduser -server $globalCatalog -filter {UserPrincipalName -eq $target} -ErrorAction "Stop"
            IF ($currentuser.count -ne 1) {Throw "$target did not return a unique value"}
            $currentUser = $currentUser[0]
            $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
            $newProxy = $currentMailbox.primarysmtpaddress.local + "@"+$onlineSMTP
            }
        Catch {
            write-error "Cannot find either the user account or mailbox for $UserName"
            return
            }
        
        #Set changes to false
        [bool]$uPNMatch = $true
        [string]$ProxyAddressUpdate = '[no update]'
        [bool]$groupsUpdated = $false

        #Check UPN to primary address
        IF ($currentUser.UserPrincipalName -ne [string]$currentMailbox.primarysmtpaddress) {
            Write-Warning "the UPN and primary SMTP do not match.  Please correct"
            [bool]$uPNMatch = $false
        }
        
        #Check for Proxy Address, add if missing
        IF ($currentMailbox.emailAddresses -notcontains $newProxy) {
            Try {
                Write-Verbose "Adding address $newProxy"
                set-mailbox $currentMailbox -Emailaddresses @{add = $newProxy}
                [string]$ProxyAddressUpdate = $newProxy
                }
            Catch {
                write-error "unable to add proxy address, check to ensure proper permissions are present!" -ErrorAction "Stop"
                }
        }

        #Check each user to be a member of the groups
        $members=@()
        ForEach ($group in $groupList) {
            try {
                $members = Get-ADGroupMember -Identity $group -server $groupDomain -Recursive | Select -ExpandProperty distinguishedname
                }
            Catch {Write-Error "cannot find group $group" -ErrorAction "Stop"}

            If ($members -notcontains $currentUser.distinguishedname) {
                Write-Verbose "adding $currentUser.Name to $group"
                Add-ADGroupMember -Identity $group -server $groupDomain -Members $currentUser.distinguishedname
                $groupsUpdated = $true
                } #End Match
            } #End ForEach

        #Need to create a custom object to add to the log
        $newentry = new-object PSObject
        $newentry | Add-Member -Type NoteProperty -Name MailboxName -Value $target
        $newentry | Add-Member -Type NoteProperty -Name UPNMatch -Value $uPNMatch
        $newentry | Add-Member -Type NoteProperty -Name ProxyAddressUpdate -Value $ProxyAddressUpdate
        $newentry | Add-Member -Type NoteProperty -Name groupsUpdated -Value $groupsUpdated
        $settingsOutLog.add($newentry) | Out-Null

        } #End ForEach

    #Save and append log
    $settingsOutLog | export-csv -Path $SettingsOutFile -Force

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

        # SettingsOutFile This specifies the filename to store batchmigration data into
        [Parameter(Mandatory=$false)] 
        [string]$SettingsOutFile = ".\MailboxSettings "+ (get-date -format m) + ".csv",

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials,

        # LocalCredentials These are the credentials require to sign into your Exchange Environment
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$LocalCredentials
    )

    #Variables specific to client
    $targetDeliveryDomain = "viacom.mail.onmicrosoft.com"
    $globalCatalog = "jumboshrimp.mtvn.ad.viacom.com:3268"

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    
    #Import UserList into a workingList
    If ($UserList) {$workingList = Get-Content $UserList}

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure they are available" -ErrorAction "Stop"}
        }

    <#
    #Load MSOnline modules
    If (!(Get-module MSOnline)) {
        Try {import-module MSOnline}
        catch {write-error "Cannot import MSOnline module, please make sure it is available" -ErrorAction "Stop"}
        }
    #>
    
    Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"

    #Connect to the Exchange online environment and track all cmdlets
    [bool]$mSOLActive = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($mSOLSession -ne $NULL) {[bool]$mSOLActive = $true}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    
    #Begin MigrationBatch
    If ($UserList) {
        
            #Set Current Session to Local Host
            IF ($mSOLActive) {
                Try {
                    $importResults = Import-PSSession $localSession -AllowClobber
                    [bool]$mSOLActive = $false
                    }
                catch {
                    write-error "can't switch context to local session" -ErrorAction "Stop"
                    }
                }

            # Set an Empty settings file before move
            [System.Collections.ArrayList]$settingsOutLog = @()

            #Set Per-User attributes on the mailbox
            Foreach ($target in $workingList) {
                Try {
                    $currentUser = @()
                    $currentUser += get-aduser -server $globalCatalog -filter {UserPrincipalName -eq $target} -ErrorAction "Stop"
                    IF ($currentuser.count -ne 1) {Throw "$target did not return a unique value"}
                    $currentUser = $currentUser[0]
                    $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
                    [string]$primarySMTP = $currentMailbox.primarysmtpaddress
                    }
                Catch {
                    write-error "Cannot find either the user account or mailbox for $target" -ErrorAction "Stop"
                    }

                #Grab current SendLimits and RetentionPolicy
                $retentionPolicy = $currentMailbox.RetentionPolicy.Name
                [bool]$updatedMBQuota = $false
                If ($currentMailbox.UseDatabaseQuotaDefaults) {
                    write-verbose "$target storage quotas are being pulled form the database, updating storage quotas"
                    $dB = get-mailboxdatabase $currentMailbox.database.name
                    IF ($dB.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {$DBReceiveQuota = $dB.ProhibitSendReceiveQuota.Value}
                    IF ($dB.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {$DBSendQuota = $dB.ProhibitSendQuota.Value}
                    IF ($dB.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {$DBWarning = $dB.IssueWarningQuota.Value}
                    set-mailbox $currentUser.Name -ProhibitSendQuota $DBSendQuota -ProhibitSendReceiveQuota $DBReceiveQuota -IssueWarningQuota $DBWarning
                    [bool]$updatedMBQuota = $true
                    }
                
                
                Write-Verbose "outputting $target settings and changes"

                #Need to create a custom object to add to the arraylist
                $newentry = new-object PSObject
                $newentry | Add-Member -Type NoteProperty -Name MailboxName -Value $target
                $newentry | Add-Member -Type NoteProperty -Name RetentionPolicy -Value $retentionPolicy
                $newentry | Add-Member -Type NoteProperty -Name UpdatedMBQuota -Value $updatedMBQuota
                $settingsOutLog.add($newentry) | Out-Null

                } # End ForEach

            #Save and append log
            $settingsOutLog | export-csv -Path $SettingsOutFile -Force

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
            Write-Verbose "Begining batch migration to O365"
            Try {
                $remoteOnboarding = "RemoteOnBoarding "+ (get-date -format m)
                $migrationEndpointOnPrem = New-MigrationEndpoint -ExchangeRemoteMove -Name OnpremEndpoint -RemoteServer $RemoteHostName -Credentials $localCredentials
                $OnboardingBatch = New-MigrationBatch -Name $remoteOnboarding -SourceEndpoint $MigrationEndpointOnprem.Identity -TargetDeliveryDomain $targetDeliveryDomain -CSVData ([System.IO.File]::ReadAllBytes($UserList))
                Start-MigrationBatch -Identity $OnboardingBatch.Identity
                } #End Try
            Catch {
                Write-Error $_.Exception.Message  
                Write-Error "MigrationBatch setup failed, review Batchname RemoteOnBoarding" -ErrorAction "Stop" 
                }
            $iD = $OnboardingBatch.Identity
            Write-Information "Migration batch $iD has been started and policies saved to $SettingsOutFile"
    } #End MigrationBatch

    #Begin Individual User
    Else {
        $target = $UserName
        #Set Current Session to Local Host
        IF ($mSOLActive) {
            Try {
                $importResults = Import-PSSession $localSession -AllowClobber
                [bool]$mSOLActive = $false
                }
            catch {
                write-error "can't switch context to local session" -ErrorAction "Stop"
                }
            }

        #Get CurrentUser and needed SMTP values
        Try {
            $currentUser = @()
            $currentUser += get-aduser -server $globalCatalog -filter {UserPrincipalName -eq $target} -ErrorAction "Stop"
            IF ($currentuser.count -ne 1) {Throw "$target did not return a unique value"}
            $currentUser = $currentUser[0]
            $currentMailbox = get-mailbox $currentUser.Name -ErrorAction "Stop"
            [string]$primarySMTP = $currentMailbox.primarysmtpaddress
            }
        Catch {
            write-error "Cannot find either the user account or mailbox for $target" -ErrorAction "Stop"
            }

        #Grab current SendLimits and RetentionPolicy
        $retentionPolicy = $currentMailbox.RetentionPolicy.Name
        If ($currentMailbox.UseDatabaseQuotaDefaults) {
            write-verbose "$target storage quotas are being pulled form the database, updating storage quotas"
            $dB = get-mailboxdatabase $currentMailbox.database.name
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
            For ($i = (Get-MoveRequestStatistics $primarySMTP).percentcomplete ; $i -lt 100 ; $i = (Get-MoveRequestStatistics $primarySMTP).percentcomplete) {
                Write-Progress -Activity "Migrating $primarySMTP, $i% complete..." -PercentComplete $i -Status "please wait"
                If ((Get-MoveRequest -Identity $primarySMTP).Status -eq "Failed" ) {throw "migration failed, see move request for more details"}
                If ((Get-MoveRequest -Identity $primarySMTP).Status -eq "Suspended" ) {throw "the job has been suspended, see move request for more details"}
                Start-Sleep -seconds 60
                } #End write-progress
            } #End Try
        Catch {
            Write-Error $_.Exception.Message  -ErrorAction "Stop" #review this one later, should contiue next on error
            }

        #Update RetentionPolicy
        Write-Verbose "Applying RetentionPolicy to $primarySMTP"
        Set-mailbox -Identity $primarySMTP -RetentionPolicy $retentionPolicy

        #Disable Clutter
        Write-Verbose "Disabling Clutter for $primarySMTP"
        Set-Clutter -Identity $primarySMTP -Enable $false
    } #End Individual User

} #End Function