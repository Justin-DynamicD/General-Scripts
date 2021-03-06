
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
        [System.Management.Automation.PSCredential]$OnlineCredentials,

        # ReportOnly This switch disables anything from being changed and only reports
        [Parameter(Mandatory=$false)]
        [switch]$ReportOnly
    )

    #Variables specific to client
    $groupList = "Office 365 Enrollment", "Office 365 Enterprise Cal" #Ensures users are a membe of listed groups.  THese have been identified as used for assigning licenses
    $groupDomain = "JohnnyKrill.corp.ad.viacom.com" #Domains that above groups are members of
    $globalCatalog = "jumboshrimp.mtvn.ad.viacom.com:3268"
    $onlineSMTP = "viacom.mail.onmicrosoft.com"
    $exchangeServer = "abfabnj50.mtvn.ad.viacom.com" #location of the Exchange cmdlets on local server
    $MSOLAccountSkuId = "viacom:ENTERPRISEPACK" #name of the license assigned to accounts

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    If ($UserList -and !(Test-Path $UserList)) {write-error "Cannot find the UserList!" -ErrorAction "Stop"}
    If ($PSVersionTable.PSVersion.Major -lt 3) {Write-Error "Powershell version is only v$($PSVersionTable.PSVersion.Major).  At least v3 must be installed" -ErrorAction "Stop"}

    #Generate Log filename
    if ($UserList -and !$SettingsOutFile) {
        $shortName = [io.path]::GetFileNameWithoutExtension($UserList)
        $SettingsOutFile = $shortName + " Init " + (get-date -format m) + ".csv"
        }
    ElseIf ($UserName -and !$SettingsOutFile) {
        $SettingsOutFile = $UserName + " Init " + (get-date -format m) + ".csv"
        }
    Write-Output "Log File set to $SettingsOutFile"

    #Import UserList into a workingList
    If ($UserList) {$workingList = Get-Content $UserList}
    Else {[array]$workingList = $UserName}

    #Load AD/MSOnline modules
    If (!(Get-module ActiveDirectory)) {
        Write-Verbose "Importing AD Module"
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure they are available" -ErrorAction "Stop"}
        }    
    If (!(Get-module MSOnline)) {
        Write-Verbose "Importing MSOnline Module"
        Try {
            import-module MSOnline
            Connect-MsolService -Credential $OnlineCredentials
            }
        catch {write-error "Cannot connect to MSOnline, please make sure the serive and modules are available" -ErrorAction "Stop"}
        }

    #Connect to the Exchange environments and track all cmdlets
    [bool]$mSOLActive = $false
    [bool]$sessionsImported = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}

    If (!$localSession) {
        Try {
            $localSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/powershell -Authentication Kerberos -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to Internal Exchange!" -ErrorAction "Stop"}
        }
    If (!$mSOLSession) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }

    # create an Empty settings log before check
    [System.Collections.ArrayList]$settingsOutLog = @()

    #Set Current Session to Office365
    Try {
        $importResults = Import-PSSession $mSOLSession -AllowClobber
        [bool]$mSOLActive = $true
        [bool]$sessionsImported = $true
        }
    catch {
        write-error "can't switch context to MSOL session" -ErrorAction "Stop"
        }
    
    #Gather list of accepted Domains for comparison
    $mSOLAcceptedDomain = Get-AcceptedDomain | select-object -ExpandProperty DomainName

    #Gather GroupMembership for each group
    [System.Collections.ArrayList]$groupMembers = @()
    ForEach ($item in $groupList) {
        Write-Output "Collecting Existing GroupMembership from $item"
        $newentry = New-Object psobject
        $newentry | Add-Member -Type NoteProperty -Name Name -Value $item
        $newentry | Add-Member -Type NoteProperty -Name Members -Value (get-adgroupmember -identity $item -server $groupDomain -recursive | Select-Object -ExpandProperty DistinguishedName)
        $groupMembers.add($newentry) | Out-Null
        } # End ForEach Loop

    #Begin per-user Loop and track progress
    [int]$totalCount = $workingList.count
    [int]$currentCount = 0
    [int]$currentPercent  = ($currentCount / $totalCount)*100
    ForEach ($target in $workingList) {
        [bool]$userExist = $true
        $currentCount ++ | Out-Null
        Write-Progress -Activity "Checking $target" -PercentComplete (($currentCount / $totalCount)*100) -Status "analyzing..."
        
        #Set Current Session to Local Host
        IF ($mSOLActive -or !$sessionsImported) {
            Write-Output "Switching to Local Session"
            Try {
                $importResults = Import-PSSession $localSession -AllowClobber
                [bool]$mSOLActive = $false
                [bool]$sessionsImported = $true
                If (!(Get-AdServerSettings).ViewEntireForest) {
                    Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"
                    }
                }
            catch {
                write-error "can't switch context to local session" -ErrorAction "Stop"
                }
            }

        #Get CurrentUser and needed SMTP values
        Write-Output "Importing settings for $target"
        Try {
            $currentUser = @()
            $currentUser += get-aduser -server $globalCatalog -filter {UserPrincipalName -eq $target} -ErrorAction "Stop"
            IF ($currentuser.count -ne 1) {Throw "$target did not return a unique value"}
            $currentUser = $currentUser[0]
            $currentMailbox = get-mailbox $currentUser.UserPrincipalName -ErrorAction "Stop"
            IF ($NULL -eq $currentMailbox) {Throw "$target mailbox could not be found locally"}
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
            Write-Output "Comparing UPN to primary SMTP"
            #Check UPN to primary address
            IF ($currentUser.UserPrincipalName -ne [string]$currentMailbox.primarysmtpaddress) {
                Write-Warning "the UPN and primary SMTP do not match.  Please correct"
                [bool]$uPNMatch = $false
                }
            
            #Check for Proxy Address, add if missing
            Write-Output "checking for existing $onlineSMTP address"
            $existingSMTPcheck = $currentMailbox.emailAddresses | Where-Object {($_ -like "smtp:*") -and ($_ -like "*@$onlineSMTP")}
            IF ($NULL -eq $existingSMTPcheck) {
                Try {
                    $newProxy = $currentMailbox.primarysmtpaddress.split("@",2)[0] +"@"+$onlineSMTP
                    Write-Output "address not found, checking if $newProxy is available."
                    If ($null -ne (get-mailbox $newProxy -ErrorAction "SilentlyContinue")) {
                        For ($i=0,($null -ne (get-mailbox $newProxy -ErrorAction "SilentlyContinue")),$i++) {
                            $newProxy = $currentMailbox.primarysmtpaddress.split("@",2)[0] + $i +"@"+$onlineSMTP
                            Write-Output "address in use, checking if $newProxy is available."
                            } #End numeric incriment
                        } #found a non-existant address!
                    Write-Output "Adding address $newProxy"
                    If (!$ReportOnly) {set-mailbox $currentMailbox.UserPrincipalName -Emailaddresses @{add = $newProxy}}
                    [string]$ProxyAddressUpdate = $newProxy
                    }
                Catch {
                    write-error "unable to add proxy address, check to ensure proper permissions are present!" -ErrorAction "Stop"
                    }
                }
            
            #Check for MSOline licenses
            Write-Output "verifying $target is synced"
            If ($NULL -eq ((Get-MsolUser -UserPrincipalName $target).ImmutableID)) {
                    [string]$mSOLLicenseUpdate = 'not a synced account'
                    write-warning "not a synced account!"
                    }
            
            If ($mSOLLicenseUpdate -ne 'not a synced account') {
                Write-Output "Importing license information"
                $isLicensed = (Get-MsolUser -UserPrincipalName $target).isLicensed
                If ($provisioningStatus = (Get-MsolUser -UserPrincipalName $target).licenses) {
                    $provisioningStatus = (Get-MsolUser -UserPrincipalName $target).licenses.servicestatus[9].provisioningstatus
                    }
                Else {$provisioningStatus = "Not Available"}
                }

            If (($provisioningStatus -ne "Success") -and ($mSOLLicenseUpdate -ne "not a synced account")) {
                Try {
                    $licenseSplat = @{UserPrincipalName = $target}
                    If (!$isLicensed) {
                        Write-Output "Adding license to import"
                        If (!$ReportOnly) {Set-MsolUser -UserPrincipalName $target -UsageLocation US}
                        $licenseSplat +=@{AddLicenses = $MSOLAccountSkuId}
                        }
                    Write-Output "Adding License Option to import"
                    $lO = New-MsolLicenseOptions -AccountSkuId $MSOLAccountSkuId
                    $licenseSplat +=@{LicenseOptions = $lO}
                    If (!$ReportOnly) {
                        Write-Output "Applying import back to account"
                        Set-MsolUserLicense @licenseSplat
                        }
                    [string]$mSOLLicenseUpdate = 'added'
                    }
                Catch {
                    write-verbose "error adding licenses to user $target"
                    [string]$mSOLLicenseUpdate = 'needed'
                    }
                }
            

            #Check each user to be a member of the groups
            ForEach ($group in $groupMembers) {
                Write-Output "checking membership of $($group.Name)"
                If ($group.members -notcontains $currentUser.distinguishedname -and !$ReportOnly) {                   
                    Write-Output "adding $($currentUser.Name) to $($group.Name)"
                    Add-ADGroupMember -Identity $group.Name -server $groupDomain -Members $currentUser    
                    $groupsUpdated = $true
                    } #End Match
                Elseif ($group.members -notcontains $currentUser.distinguishedname -and $ReportOnly) {
                    $groupsUpdated = $true
                    }
                } #End ForEach

            #Gather all smtp addresses and compare to online list
            Write-Output "comparing smtp addresses to online list"
            $addr = $currentMailbox.emailaddresses | Where-Object {$_ -like "smtp:*"}
            $addr = ($addr | foreach-Object {($_.split("@",2))[1]})
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
    Write-Output "dumping log file"
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
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$LocalCredentials
    )

    #Variables specific to client
    $targetDeliveryDomain = "viacom.mail.onmicrosoft.com"
    $globalCatalog = "jumboshrimp.mtvn.ad.viacom.com:3268"
    $exchangeServer = "abfabnj50.mtvn.ad.viacom.com" #location of the local server

    #Validate parameter combinations are valid
    If ($UserName -and $UserList) {write-error "You can only specify either UserName or UserList, not both" -ErrorAction "Stop"}
    If (!$UserName -and !$UserList) {write-error "You must specify either UserName or UserList" -ErrorAction "Stop"}
    If ($UserList -and !(Test-Path $UserList)) {write-error "Cannot find the UserList!" -ErrorAction "Stop"}
    if (($NULL -eq $LocalCredentials) -and ($UserName)) {write-error "O365 requires -LocalCredentials to be supplied in order to move single mailboxes" -ErrorAction "Stop"}
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
    [bool]$sessionsImported = $false
    [bool]$mSOLActive = $false
    $localSession = Get-PSSession | Where-Object {$_.ComputerName -like "*.viacom.com"}
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($NULL -ne $mSOLSession) {[bool]$mSOLActive = $true}

    If (!$localSession) {
        Try {
            $localSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeServer/powershell -Authentication Kerberos -AllowRedirection
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
    
    #Begin MigrationBatch
    If ($UserList) {
        
            #Set Current Session to Local Host
            IF ($mSOLActive -or !$sessionsImported) {
                Try {
                    $importResults = Import-PSSession $localSession -AllowClobber
                    [bool]$sessionsImported = $true
                    [bool]$mSOLActive = $false
                     If (!(Get-AdServerSettings).ViewEntireForest) {
                        Set-ADServerSettings -ViewEntireForest $true
                        }
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
                    IF ($NULL -eq $currentMailbox) {Throw "$target mailbox could not be found locally"}
                    [string]$primarySMTP = $currentMailbox.primarysmtpaddress
                    }
                Catch {
                    write-error "Cannot find either the user account or mailbox for $target" -ErrorAction "Stop"
                    }

                #Grab current SendLimits and RetentionPolicy
                $retentionPolicy = $currentMailbox.RetentionPolicy
                If ($currentMailbox.UseDatabaseQuotaDefaults) {
                    write-verbose "$target storage quotas are being pulled form the database, updating storage quotas"
                    $dB = get-mailboxdatabase $currentMailbox.database
                    $DBReceiveQuota = $dB.ProhibitSendReceiveQuota
                    $DBSendQuota = $dB.ProhibitSendQuota
                    $DBWarning = $dB.IssueWarningQuota
                    }
                Else {
                    $DBReceiveQuota = $currentMailbox.ProhibitSendReceiveQuota
                    $DBSendQuota = $currentMailbox.ProhibitSendQuota
                    $DBWarning = $currentMailbox.IssueWarningQuota
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
                    [bool]$sessionsImported = $true
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
                Write-Error "MigrationBatch setup failed, review Batchname $remoteOnBoarding" -ErrorAction "Stop" 
                }
            
            Write-Output " "
            Write-Output "Migration batch $remoteOnboarding has been started and policies saved to $SettingsOutFile"
            Remove-Item -Path $tmpFileName -force

    } #End MigrationBatch

    #Begin Individual User
    Else {
        $target = $UserName
        #Set Current Session to Local Host
        IF ($mSOLActive -or !$sessionsImported) {
            Try {
                $importResults = Import-PSSession $localSession -AllowClobber
                [bool]$sessionsImported = $true
                [bool]$mSOLActive = $false
                 If (!(Get-AdServerSettings).ViewEntireForest) {
                    Set-ADServerSettings -ViewEntireForest $true -WarningAction "SilentlyContinue"
                    }
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
            IF ($NULL -eq $currentMailbox) {Throw "$target mailbox could not be found locally"}
            [string]$primarySMTP = $currentMailbox.primarysmtpaddress
            }
        Catch {
            write-error "Cannot find either the user account or mailbox for $target" -ErrorAction "Stop"
            }

        #Grab current SendLimits and RetentionPolicy
        $retentionPolicy = $currentMailbox.RetentionPolicy
        If ($currentMailbox.UseDatabaseQuotaDefaults) {
            write-verbose "$target storage quotas are being pulled form the database, updating storage quotas"
            $dB = get-mailboxdatabase $currentMailbox.database
            $DBReceiveQuota = $dB.ProhibitSendReceiveQuota
            $DBSendQuota = $dB.ProhibitSendQuota
            $DBWarning = $dB.IssueWarningQuota
            }
        Else {
            $DBReceiveQuota =$currentMailbox.ProhibitSendReceiveQuota
            $DBSendQuota = $currentMailbox.ProhibitSendQuota
            $DBWarning = $currentMailbox.IssueWarningQuota
            }

        #switch session to mSOL
        IF (!$mSOLActive) {
            Try {
                $importResults = Import-PSSession $mSOLSession -AllowClobber
                Write-Verbose $importResults
                [bool]$sessionsImported = $true
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
            Write-Error $_.Exception.Message  -ErrorAction "Stop"
            }
        
        #Define a hashtable to splat settings
        $splatMailbox = @{Identity = $primarySMTP}

        #update Storage Policy
        If ($DBSendQuota -ne "unlimited") {$DBSendQuota = ($currentUser.DBSendQuota).split(" ",3)[0] + ($currentUser.DBSendQuota).split(" ",3)[1]}
        If ($DBReceiveQuota -ne "unlimited") {$DBReceiveQuota = ($DBReceiveQuota).split(" ",3)[0] + ($DBReceiveQuota).split(" ",3)[1]}
        If ($DBWarning -ne "unlimited") {$DBWarning = ($DBWarning).split(" ",3)[0] + ($DBWarning).split(" ",3)[1]}
        $splatMailbox += @{ProhibitSendQuota = $DBSendQuota}
        $splatMailbox += @{ProhibitSendReceiveQuota = $DBReceiveQuota}
        $splatMailbox += @{IssueWarningQuota = $DBWarning}

        #Update RetentionPolicy and Deleteditemretention
        $splatMailbox += @{RetentionPolicy = $retentionPolicy}
        $splatMailbox += @{RetainDeletedItemsFor = 30}
        
        #Apply Settings to mailbox
        write-verbose "Updating Mailbox with stored settings"
        Set-mailbox $splatMailbox

        #Disable Clutter
        Write-Verbose "Disabling Clutter for $primarySMTP"
        Set-Clutter -Identity $primarySMTP -Enable $false
    } #End Individual User

} #End Function

function Complete-O365User 
{
    <#
    .Synopsis
    This function is used in order to complete migration batches and importing/applying te settings captured by move-O365User in batchmode.
    .DESCRIPTION
    This function is used in order to complete migration batches and importing/applying te settings captured by move-O365User in batchmode.
    It only needs to be run with batch migrations, as individual moves will complete automatically.
    .EXAMPLE
    Complete-O365User -MigrationBatch "Migration Batch GroupC" -SettingsOutFile "Migration Batch GroupC Logs.csv" -OnlineCredentials (Get-Credentials)
    .NOTES
    General notes
    #>

    param (
        # This is the Migration batch you wish to complete
        [Parameter(Mandatory=$false)] 
        [string]$MigrationBatch,

        # SettingsOutFile This specifies the filename to pull restore data from
        [Parameter(Mandatory=$true)] 
        [string]$SettingsOutFile,

        # OnlineCredentials These are the credentials required to sign into your O365 tenant
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$OnlineCredentials
    )

    #Variables specific to client
    $groupList = "zdm Migration to Office 365" #Ensures users are a member of listed groups.
    $groupDomain = "paramount.ad.viacom.com" #Domains that above groups are members of

    #Validate parameter combinations are valid
    If (!(Test-Path $SettingsOutFile)) {write-error "Cannot find $SettingsOutFile" -ErrorAction "Stop"}
    If ($PSVersionTable.PSVersion.Major -lt 3) {Write-Error "Powershell version is only v$($PSVersionTable.PSVersion.Major).  At least v3 must be installed" -ErrorAction "Stop"}

    #Load AD/MSOnline modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirectory module, please make sure it is available" -ErrorAction "Stop"}
        }    

    #Generate Migration Batch name if not provided
    if (!$MigrationBatch) {
        $shortName = [io.path]::GetFileNameWithoutExtension($SettingsOutFile)
        $MigrationBatch = ($shortName.split(" ",2))[0]
        }

    #Import UserList into a workingList
    $workingList = import-csv $SettingsOutFile

    #Connect to the Exchange online environment and track all cmdlets
    [bool]$mSOLActive = $false
    [bool]$sessionsImported = $false
    $mSOLSession = Get-PSSession | Where-Object {$_.ComputerName -eq "ps.outlook.com"}
    If ($NULL -ne $mSOLSession) {[bool]$mSOLActive = $true}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }
    
    $mSOLModule = Import-PSSession $mSOLSession -AllowClobber
    [bool]$mSOLActive = $true
    [bool]$sessionsImported = $true

    #Look for MigrationBatch and Complete Migration
    $currentBatch = Get-MigrationBatch $MigrationBatch
    If ($NULL -eq $currentBatch) {
        write-error "cannot find Migration Batch $MigrationBatch, please verify it exists or use -MigrationBatch flag" -ErrorAction "Stop"
        }
    Else {
        If ($currentBatch.Status.value -like "Synced*") {
            Complete-MigrationBatch -Identity $currentBatch.Identity.Name
            while ((Get-MigrationBatch $MigrationBatch).Status.value -notlike "Completed*") {
                Write-Output "$MigrationBatch is $((Get-MigrationBatch $MigrationBatch).Status.value), sleeping for 60 seconds..."
                Start-Sleep -seconds 60
                } #End Wait
            }
        ElseIf ($currentBatch.Status.value -notlike "Completed*") {
            Write-Error -Message "Batch $MigrationBatch is not in a synced or completed Status, cannot continue" -ErrorAction "Stop"
            }
        }

    #Gather GroupMembership for each group
    [System.Collections.ArrayList]$groupMembers = @()
    ForEach ($item in $groupList) {
        Write-Output "Collecting Existing GroupMembership from $item"
        $newentry = New-Object psobject
        $newentry | Add-Member -Type NoteProperty -Name Name -Value $item
        $newentry | Add-Member -Type NoteProperty -Name Members -Value (get-adgroupmember -identity $item -server $groupDomain -recursive | Select-Object -ExpandProperty DistinguishedName)
        $groupMembers.add($newentry) | Out-Null
        } # End ForEach Loop

    #begin user processing to reapply settings
    [int]$totalCount = $workingList.count
    [int]$currentCount = 0
    foreach ($currentUser in $workingList) {
        $currentCount ++ | Out-Null
        Write-Progress -Activity "Updating $($currentUser.MailboxName)" -PercentComplete (($currentCount / $totalCount)*100) -Status "updating..."
        
        #Define a hashtable to splat settings
        Write-Output "Creating Settings table for $($currentUser.MailboxName)"
        $splatMailbox = @{Identity = $currentUser.MailboxName}   

        #update Storage Policy
        $DBSendQuota = ($currentUser.DBSendQuota).split(" ",3)[0] + ($currentUser.DBSendQuota).split(" ",3)[1]
        $DBReceiveQuota = ($currentUser.DBReceiveQuota).split(" ",3)[0] + ($currentUser.DBReceiveQuota).split(" ",3)[1]
        $DBWarning = ($currentUser.DBWarning).split(" ",3)[0] + ($currentUser.DBWarning).split(" ",3)[1]
        Write-Output "Adding ProhibitSendQuota = $DBSendQuota"
        $splatMailbox += @{ProhibitSendQuota = $DBSendQuota}
        Write-Output "Adding ProhibitSendReceiveQuota = $DBReceiveQuota"
        $splatMailbox += @{ProhibitSendReceiveQuota = $DBReceiveQuota}
        Write-Output "Adding IssueWarningQuota = $DBWarning"
        $splatMailbox += @{IssueWarningQuota = $DBWarning}

        #Update RetentionPolicy and Deleteditemretention
        Write-Output "Adding RetentionPolicy = $($currentUser.retentionPolicy)"
        $splatMailbox += @{RetentionPolicy = $currentUser.retentionPolicy}
        Write-Output "Adding RetainDeletedItemsFor = 30"
        $splatMailbox += @{RetainDeletedItemsFor = 30}

        #Apply Settings to mailbox
        Write-Output "Updating Mailbox with stored settings for $($currentUser.MailboxName)"
        Set-mailbox $splatMailbox

        #Disable Clutter
        Write-Output "Disabling Clutter for $($currentUser.MailboxName)"
        Set-Clutter -Identity $currentUser.MailboxName -Enable $false | Out-Null

        #Check each user to be a member of the groups if they are part of the group's domain
        $isMember =  get-aduser -server $groupDomain -filter {UserPrincipalName -eq $currentUser.MailboxName}
        If ($isMember) {
            Write-Output "$($currentUser.MailboxName) is a member of the domain $groupDomain"
            ForEach ($group in $groupMembers) {
                Write-Output "checking membership of $($group.Name)"
                If ($group.members -notcontains $currentUser.distinguishedname) {                   
                    Write-Output "adding $($currentUser.Name) to $($group.Name)"
                    Add-ADGroupMember -Identity $group.Name -server $groupDomain -Members $currentUser    
                    $groupsUpdated = $true
                    } #End Match
                } #End ForEach
            }
        Else {
            Write-Output "$($currentUser.MailboxName) is not a member of the domain $groupDomain"
            $groupsUpdated = $false
            }
        } #End CurrentUser-ForEach

} #End Function
