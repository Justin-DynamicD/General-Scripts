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
        [String]$UserName
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
    $currentUser = get-aduser -filter {name -eq $UserName}
    $currentMailbox = get-mailbox $UserName
    $newProxy = "smtp:"+$UserName+"@"+$onlineSMTP

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

    #Get CurrentUser and needed SMTP values
    Try {
        $currentUser = get-aduser -filter {name -eq $UserName} -ErrorAction "Stop"
        $currentMailbox = get-mailbox $UserName -ErrorAction "Stop"
        $primarySMTP = $currentMailbox.primarysmtpaddress
        }
    Catch {
        write-error "Cannot find either the user account or mailbox for $UserName" -ErrorAction "Stop"
        }

    #Connect to the Exchange online environment and clobber all modules
    Try {
        [bool]$mSOLActive = $true
        Get-PSSession ps.outlook.com -ErrorAction "Stop"
        }
    Catch {[bool]$mSOLActive = $false}

    If (!$mSOLActive) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $OnlineCredentials -Authentication Basic -AllowRedirection
            $importResults = Import-PSSession $mSOLSession -AllowClobber
            Write-Verbose $importResults
            } #End Try
        Catch {write-error "Cannot connect to O365" -ErrorAction "Stop"}
        }  
    
    #Variables specific to client
    $targetDeliveryDomain = "viacom.mail.onmicrosoft.com"

    # Do the move
    New-MoveRequest -Identity $primarySMTP -Remote -RemoteHostName $RemoteHostName -RemoteCredential $LocalCredentials -TargetDeliveryDomain $targetDeliveryDomain -BadItemLimit 100 -AcceptLargeDataLoss

} #End Function