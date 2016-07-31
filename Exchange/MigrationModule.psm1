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
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias('user')] 
        [String]$UserName,

        # Credentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias('creds')]
        [PsCredential]$Credentials
    
    )

    #Variables specific to client
    $groupList = "Office 365 Enrollment", "Office 365 Enterprise Cal"

    #Load AD modules
    If (!(Get-module ActiveDirectory)) {
        Try {import-module ActiveDirectory}
        catch {write-error "Cannot import ActiveDirecotry modules, please make sure htey are avialable" -ErrorAction "Stop"}
        }
    
    #Connect to the Exchange online environment and clobber all modules
    If (!(Get-module mSOLSession)) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Credentials -Authentication Basic -AllowRedirection
            $importResults = Import-PSSession $mSOLSession -AllowClobber
            Write-Verbose $importResults
            } #End Try
        Catch {write-error "Cannot import MSOL modules, please make sure they are avialable" -ErrorAction "Stop"}
        }

    #Check each user to be a member of the groups
    $members=@()
    [bool]$isUpdated = $false
    ForEach ($group in $groupList) {
        $members = Get-ADGroupMember -Identity $group -Recursive | Select -ExpandProperty Name
        If ($members -notcontains $UserName) {
            Write-Verbose "adding $UserName to $group"
            Add-ADGroupMember -Identity $group -Members $UserName
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
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias('user')] 
        [String]$UserName,

        # Credentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias('creds')]
        [PsCredential]$Credentials
    )

    #Connect to the Exchange online environment and clobber all modules
    If (!(Get-module mSOLSession)) {
        Try {
            $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Credentials -Authentication Basic -AllowRedirection
            $importResults = Import-PSSession $mSOLSession -AllowClobber
            Write-Verbose $importResults
            } #End Try
        Catch {write-error "Cannot import MSOL modules, please make sure they are avialable" -ErrorAction "Stop"}
        }  

} #End Function