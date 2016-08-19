
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
        [string]$SettingsOutFile,

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials

    )

    #Variables specific to client
    $groupList = "Office 365 Enrollment", "Office 365 Enterprise Cal" #Ensures users are a membe of listed groups.  THese have been identified as used for assigning licenses
    $groupDomain = "corp.ad.viacom.com" #Domains that above groups are members of
    $globalCatalog = "jumboshrimp.mtvn.ad.viacom.com:3268"
    $onlineSMTP = "viacom.mail.onmicrosoft.com"
    $exchangeModules = "E:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1" #location of the Exchange cmdlets on local server

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    If ($UserList -and !(Test-Path $UserList)) {write-error "Cannot find the UserList!" -ErrorAction "Stop"}
    If ($PSVersionTable.PSVersion.Major -lt 3) {Write-Error "Powershell version is only $($PSVersionTable.PSVersion.Major).  At least 3 must be installed" -ErrorAction "Stop"}

    #Generate Log filename
    if ($UserList -and !$SettingsOutFile) {
        $shortName = [io.path]::GetFileNameWithoutExtension($UserList)
        $SettingsOutFile = $shortName + " Log " + (get-date -format m) + ".csv"
        }
    ElseIf ($UserName -and !$SettingsOutFile) {
        $SettingsOutFile = $UserName + " Log " + (get-date -format m) + ".csv"
        }
    
    #Import UserList into a workingList
    If ($UserList) {$workingList = Get-Content $UserList}
    Else {[array]$workingList = $UserName}

    #Load AD/MSOnline modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure they are available" -ErrorAction "Stop"}
        }    
    If (!(Get-module MSOnline)) {
        Try {
            import-module MSOnline
            Connect-MsolService -Credential $OnlineCredentials
            }
        catch {write-error "Cannot connect to MSOnline, please make sure the serive and modules are available" -ErrorAction "Stop"}
        }

    #Connect to the Exchange environments and track all cmdlets
    [bool]$mSOLActive = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}

    If (!$localSession) {
        Try {
            . $exchangeModules
            Connect-ExchangeServer -Auto
            $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
            } #End Try
        Catch {write-error "Cannot connect to INternal Exchange!" -ErrorAction "Stop"}
        }
    If (!$mSOLSession) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }

    If (!(Get-AdServerSettings).ViewEntireForest) {
        Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"
        }

    # create an Empty settings log before check
    [System.Collections.ArrayList]$settingsOutLog = @()

    #Set Current Session to Office365
    Try {
        $importResults = Import-PSSession $mSOLSession -AllowClobber
        [bool]$mSOLActive = $true
        }
    catch {
        write-error "can't switch context to MSOL session" -ErrorAction "Stop"
        }
    
    #Gather list of accepted Domains for comparison
    $mSOLAcceptedDomain = Get-AcceptedDomain | select -ExpandProperty DomainName

    #Begin per-user Loop and track progress
    [int]$totalCount = $workingList.count
    [int]$currentCount = 0
    [int]$currentPercent  = ($currentCount / $totalCount)*100
    ForEach ($target in $workingList) {
        [bool]$userExist = $true
        $currentCount ++ | Out-Null
        Write-Progress -Activity "Checking $target" -PercentComplete (($currentCount / $totalCount)*100) -Status "analyzing..."
        
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
            $currentMailbox = get-mailbox $currentUser.UserPrincipalName -ErrorAction "Stop"
            IF ($currentMailbox -eq $NULL) {Throw "$target mailbox could not be found locally"}
            }
        Catch {
            [bool]$userExist = $false
            write-error "Cannot find either the user account or mailbox for $target" -ErrorAction "SilentlyContinue"
            }
        
        #Set changes to false
        [bool]$uPNMatch = $true
        [string]$ProxyAddressUpdate = '[no update]'
        [bool]$groupsUpdated = $false
        [string]$mSOLLicenseUpdate = ''

        #UserExist Check
        If ($userExist) {

            #Check UPN to primary address
            IF ($currentUser.UserPrincipalName -ne [string]$currentMailbox.primarysmtpaddress) {
                Write-Warning "the UPN and primary SMTP do not match.  Please correct"
                [bool]$uPNMatch = $false
                }
            
            #Check for Proxy Address, add if missing
            $existingSMTPcheck = $currentMailbox.emailAddresses | where {($_.PrefixString -eq "smtp") -and ($_.AddressString -like "*@$onlineSMTP")}
            IF ($existingSMTPcheck -eq $NULL) {
                Try {
                    $newProxy = $currentMailbox.primarysmtpaddress.local +"@"+$onlineSMTP
                    If ((get-mailbox $newProxy) -ne $null) {
                        For ($i=0,((get-mailbox $newProxy) -ne $null),$i++) {
                            $newProxy = $currentMailbox.primarysmtpaddress.local + $i +"@"+$onlineSMTP
                            } #End numeric incriment
                        } #found a non-existant address!
                    Write-Verbose "Adding address $newProxy"
                    set-mailbox $currentMailbox -Emailaddresses @{add = $newProxy}
                    [string]$ProxyAddressUpdate = $newProxy
                    }
                Catch {
                    write-error "unable to add proxy address, check to ensure proper permissions are present!" -ErrorAction "Stop"
                    }
                }
            
            #Check for MSOline licenses
            If (!(Get-MsolUser -UserPrincipalName $target).isLicensed) {
                Try {
                    Set-MsolUser -UserPrincipalName $target -UsageLocation US
                    Set-MsolUserLicense -UserPrincipalName $target -AddLicenses viacom:ENTERPRISEPACK
                    [string]$mSOLLicenseUpdate = 'added'
                    }
                Catch {
                    write-verbose "error adding licenses to user $target"
                    [string]$mSOLLicenseUpdate = 'needed'
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
                    Add-ADGroupMember -Identity $group -server $groupDomain -Members $currentUser
                    $groupsUpdated = $true
                    } #End Match
                } #End ForEach

            #Gather all smtp addresses and compare to online list
            $addr = $currentMailbox.emailaddresses | Select -ExpandProperty ProxyAddressString | Where-Object {$_ -like "smtp:*"}
            $addr = ($addr | foreach {($_.split("@",2))[1]})
            [System.Collections.ArrayList]$missingDomainList = @()
            Foreach ($item in $addr) {If ($mSOLAcceptedDomain -notcontains $item) {$missingDomainList.add($item) | Out-Null }}
            [string]$missingDomainList = $missingDomainList -join "`r`n"
            
            } # End If userExist
        
        Else {
            $uPNMatch = $false
            $ProxyAddressUpdate = '[user not found]'
            $groupsUpdated = $false
            $missingDomainList = '[user not found]'
            $mSOLLicenseUpdate = '[user not found]'
            }

        #Need to create a custom object to add to the log
        $newentry = new-object PSObject
        $newentry | Add-Member -Type NoteProperty -Name MailboxName -Value $target
        $newentry | Add-Member -Type NoteProperty -Name UPNMatch -Value $uPNMatch
        $newentry | Add-Member -Type NoteProperty -Name ProxyAddressUpdate -Value $ProxyAddressUpdate
        $newentry | Add-Member -Type NoteProperty -Name groupsUpdated -Value $groupsUpdated
        $newentry | Add-Member -Type NoteProperty -Name mSOLLicenseUpdate -Value $mSOLLicenseUpdate
        $newentry | Add-Member -Type NoteProperty -Name MissingDomains -Value $missingDomainList
        $settingsOutLog.add($newentry) | Out-Null

        } #End ForEach

    #Save and append log
    $settingsOutLog | export-csv -Path $SettingsOutFile -Force

} #End Function


function Move-O365User 
{
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
        [string]$SettingsOutFile,

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
    $exchangeModules = "E:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1" #location of the Exchange cmdlets on local server

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    If ($UserList -and !(Test-Path $UserList)) {write-error "Cannot find the UserList!" -ErrorAction "Stop"}
    If ($PSVersionTable.PSVersion.Major -lt 3) {Write-Error "Powershell version is only $($PSVersionTable.PSVersion.Major).  At least 3 must be installed" -ErrorAction "Stop"}
    
    #Generate Log filename
    if ($UserList -and !$SettingsOutFile) {
        $shortName = [io.path]::GetFileNameWithoutExtension($UserList)
        $SettingsOutFile = $shortName + " Log " + (get-date -format m) + ".csv"
        }
    ElseIf ($UserName -and !$SettingsOutFile) {
        $SettingsOutFile = $UserName + " Log " + (get-date -format m) + ".csv"
        }

    #Import UserList into a workingList
    If ($UserList) {$workingList = Get-Content $UserList}

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure they are available" -ErrorAction "Stop"}
        }

    #Connect to the Exchange online environment and track all cmdlets
    [bool]$mSOLActive = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($mSOLSession -ne $NULL) {[bool]$mSOLActive = $true}


    If (!$localSession) {
        Try {
            . $exchangeModules
            Connect-ExchangeServer -Auto
            $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
            [bool]$mSOLActive = $false
            } #End Try
        Catch {write-error "Cannot connect to Internal Exchange!" -ErrorAction "Stop"}
        }
    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    
    If (!(Get-AdServerSettings).ViewEntireForest) {
        Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"
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
                    $currentMailbox = get-mailbox $currentUser.UserPrincipalName -ErrorAction "Stop"
                    IF ($currentMailbox -eq $NULL) {Throw "$target mailbox could not be found locally"}
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
                    IF ($dB.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {[string]$DBReceiveQuota = $dB.ProhibitSendReceiveQuota.Value}
                    IF ($dB.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {[string]$DBSendQuota = $dB.ProhibitSendQuota.Value}
                    IF ($dB.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {[string]$DBWarning = $dB.IssueWarningQuota.Value}
                    }
                Else {
                    IF ($currentMailbox.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {[string]$DBReceiveQuota = $currentMailbox.ProhibitSendReceiveQuota.Value}
                    IF ($currentMailbox.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {[string]$DBSendQuota = $currentMailbox.ProhibitSendQuota.Value}
                    IF ($currentMailbox.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {[string]$DBWarning = $currentMailbox.IssueWarningQuota.Value}
                    }
                
                
                Write-Verbose "outputting $target settings and changes"

                #Need to create a custom object to add to the arraylist
                $newentry = new-object PSObject
                $newentry | Add-Member -Type NoteProperty -Name MailboxName -Value $target
                $newentry | Add-Member -Type NoteProperty -Name RetentionPolicy -Value $retentionPolicy
                $newentry | Add-Member -Type NoteProperty -Name DBReceiveQuota -Value $DBReceiveQuota
                $newentry | Add-Member -Type NoteProperty -Name DBSendQuota -Value $DBSendQuota
                $newentry | Add-Member -Type NoteProperty -Name DBWarning -Value $DBWarning
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

                #Create new import file that matches format required by batchmigration
                $rSet=$Null
                $tmpFileName=$Null
                $set = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
                for ($i=1; $i -le 6; $i++) {
                    $rSet += $set | Get-Random
                    }
                $tmpFileName = $env:temp + "\" + $Rset + ".csv"
                "emailaddress" | Out-File $tmpFileName -Force
                $workingList | Out-File $tmpFileName -Append

                #Name and start the batch
                $shortName = [io.path]::GetFileNameWithoutExtension($UserList)
                $remoteOnboarding = $shortName + " " + (get-date -format m)
                New-MigrationBatch -Name $remoteOnboarding -SourceEndpoint $RemoteHostName -TargetDeliveryDomain $targetDeliveryDomain -CSVData ([System.IO.File]::ReadAllBytes($tmpFileName)) -baditemlimit 100 -autostart
                } #End Try
            Catch {
                Write-Error $_.Exception.Message  
                Write-Error "MigrationBatch setup failed, review Batchname RemoteOnBoarding" -ErrorAction "Stop" 
                }

            Write-Output "Migration batch $remoteOnboarding has been started and policies saved to $SettingsOutFile"
            Remove-Item -Path $tmpFileName -force

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
            $currentMailbox = get-mailbox $currentUser.UserPrincipalName -ErrorAction "Stop"
            IF ($currentMailbox -eq $NULL) {Throw "$target mailbox could not be found locally"}
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
            }
        Else {
            IF ($currentMailbox.ProhibitSendReceiveQuota.IsUnlimited) {$DBReceiveQuota = "Unlimited"} Else {$DBReceiveQuota = $currentMailbox.ProhibitSendReceiveQuota.Value}
            IF ($currentMailbox.ProhibitSendQuota.IsUnlimited) {$DBSendQuota = "Unlimited"} Else {$DBSendQuota = $currentMailbox.ProhibitSendQuota.Value}
            IF ($currentMailbox.IssueWarningQuota.IsUnlimited) {$DBWarning = "Unlimited"} Else {$DBWarning = $currentMailbox.IssueWarningQuota.Value}
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

        #update Storage Policy
        set-mailbox $currentUser.Name -UseDatabaseQuotaDefaults $false -ProhibitSendQuota $DBSendQuota -ProhibitSendReceiveQuota $DBReceiveQuota -IssueWarningQuota $DBWarning

        #Update RetentionPolicy
        Write-Verbose "Applying RetentionPolicy to $primarySMTP"
        Set-mailbox -Identity $primarySMTP -RetentionPolicy $retentionPolicy

        #Disable Clutter
        Write-Verbose "Disabling Clutter for $primarySMTP"
        Set-Clutter -Identity $primarySMTP -Enable $false
    } #End Individual User

} #End Function

function Complete-O365User 
{
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
    .FUNCTIONALITY
    The functionality that best describes this cmdlet
    #>

    param (
        # This is the Migration batch you wish to complete
        [Parameter(Mandatory=$false)] 
        [string]$MigrationBatch,

        # SettingsOutFile This specifies the filename to store batchmigration data into
        [Parameter(Mandatory=$true)] 
        [string]$SettingsOutFile,

        # OnlineCredentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials
    )

    #Variables specific to client
    $targetDeliveryDomain = "viacom.mail.onmicrosoft.com"
    $globalCatalog = "jumboshrimp.mtvn.ad.viacom.com:3268"
    

    #Validate parameter combinations are valid
    If (!(Test-Path $SettingsOutFile)) {write-error "Cannot find $SettingsOutFile" -ErrorAction "Stop"}
    If ($PSVersionTable.PSVersion.Major -lt 3) {Write-Error "Powershell version is only $($PSVersionTable.PSVersion.Major).  At least 3 must be installed" -ErrorAction "Stop"}

    #Generate Migration Batch name if not provided
    if (!$MigrationBatch) {
        $shortName = [io.path]::GetFileNameWithoutExtension($SettingsOutFile)
        $MigrationBatch = ($shortName.split(" ",2))[0]
        }

    #Import UserList into a workingList
    $workingList = import-csv $SettingsOutFile

    #Connect to the Exchange online environment and track all cmdlets
    [bool]$mSOLActive = $false
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($mSOLSession -ne $NULL) {[bool]$mSOLActive = $true}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    
    Import-PSSession $mSOLSession -AllowClobber
    [bool]$mSOLActive = $true

    #Look for MigrationBatch and Complete Migration
    $currentBatch = Get-MigrationBatch $MigrationBatch
    If ($currentBatch -eq $NULL) {
        write-error "cannot find Migration Batch $MigrationBatch, please verify it exists or use -MigrationBatch flag" -ErrorAction "Stop"
        }
    Else {
        If ($currentBatch.Status.value -like "Synced*") {
            Complete-MigrationBatch -Identity $currentBatch.Identity -force
            while ((Get-MigrationBatch $currentBatch).Status.value -ne "Completed") {Start-Sleep -seconds 60}
            }
        }

    #begin user processing to reapply settings
    [int]$totalCount = $workingList.count
    [int]$currentCount = 0
    foreach ($currentUser in $workingList) {
        $currentCount ++ | Out-Null
        Write-Progress -Activity "Checking $currentUser" -PercentComplete (($currentCount / $totalCount)*100) -Status "updating..."
        
        #update Storage Policy
        Write-Verbose "Applying StorageQuotas to $($currentUser.MailboxName)"
        $DBSendQuota = ($currentUser.DBSendQuota).split(" ",3)[0] + ($currentUser.DBSendQuota).split(" ",3)[1]
        $DBReceiveQuota = ($currentUser.DBReceiveQuota).split(" ",3)[0] + ($currentUser.DBReceiveQuota).split(" ",3)[1]
        $DBWarning = ($currentUser.DBWarning).split(" ",3)[0] + ($currentUser.DBWarning).split(" ",3)[1]

        set-mailbox $currentUser.MailboxName -UseDatabaseQuotaDefaults $false -ProhibitSendQuota $DBSendQuota -ProhibitSendReceiveQuota $DBReceiveQuota -IssueWarningQuota $DBWarning

        #Update RetentionPolicy
        Write-Verbose "Applying RetentionPolicy to $($currentUser.MailboxName)"
        Set-mailbox -Identity $currentUser.MailboxName -RetentionPolicy $currentUser.retentionPolicy

        #Disable Clutter
        Write-Verbose "Disabling Clutter for $($currentUser.MailboxName)"
        Set-Clutter -Identity $currentUser.MailboxName -Enable $false

        } #End ForEach

} #End Function

function New-ManagedFolder
{
    Param(
        [Parameter(Mandatory=$True)]
            [string]$TargetMailbox,
        [Parameter(Mandatory=$True)]
            [string]$FolderName,
        [Parameter(Mandatory=$True)]
            [string]$RetentionTag,
        [Parameter(Mandatory=$False)]
            [string]$AutoD = $True,
        [Parameter(Mandatory=$False)]
            [string]$EwsUri = "https://mail.office365.com/ews/exchange.asmx",
        [Parameter(Mandatory=$False)]
            [string]$ApiPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
        [Parameter(Mandatory=$False)]
            [string]$Version = "Exchange2013_SP1"
    )

    $ImpersonationCreds = Get-Credential -Message "Enter Credentials for Account with Impersonation Role..."

    Add-Type -Path $ApiPath

    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$Version
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

    $Creds = New-Object System.Net.NetworkCredential($ImpersonationCreds.UserName, $ImpersonationCreds.Password)   
    $Service.Credentials = $Creds

    if ($AutoD -eq $True) {
        $Service.AutodiscoverUrl($TargetMailbox,{$True})  
        "EWS URI = " + $Service.url
    }
    else {
        $Uri=[system.URI] $EwsUri
        $Service.Url = $uri
    }

    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)

    $Folder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)  
    $Folder.DisplayName = $FolderName
    $Folder.FolderClass = "IPF.Note"

    $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$TargetMailbox)   
    $EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$FolderId)

    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)  
    $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)  
    $FindFolderResults = $Service.FindFolders($EWSParentFolder.Id,$SearchFilter,$FolderView)

    if ($FindFolderResults.TotalCount -eq 0) {  

        $Tag = ($Service.GetUserRetentionPolicyTags().RetentionPolicyTags | where {$_.DisplayName -eq $RetentionTag})

        $Folder.PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.PolicyTag($true,$Tag.RetentionId)
        $Folder.Save($EWSParentFolder.Id)
    }  
    elseif ($FindFolderResults.TotalCount -eq 1) {  
        Write-Verbose ("The folder '$FolderName' already exists in mailbox '$TargetMailbox'")
        $Tag = ($Service.GetUserRetentionPolicyTags().RetentionPolicyTags | where {$_.DisplayName -eq $RetentionTag})
        $Folder = $FindFolderResults[0]
        $Folder.PolicyTag = New-Object Microsoft.Exchange.WebServices.Data.PolicyTag($true,$Tag.RetentionId)
        $Folder.Save($EWSParentFolder.Id)
    }
    else {
        Write-Verbose "found multiple instances of the desired folder"
    }
}