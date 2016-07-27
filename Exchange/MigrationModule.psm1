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
    [CmdletBinding(SupportsShouldProcess=$true, PositionalBinding=$true, ConfirmImpact='Medium')]
    [OutputType([String])]

    Param
    (
        # UserName this is the name of the account you wish to prepare for migration
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias('user')] 
        [String]$UserName,

        # Credentials These are the credentials require to sign into your O365 tenant
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias('creds')]
        [PsCredential]$Credentials
    
    )

    Begin
    {
        #Connect to the Exchagne online environment and clobber
        $mSOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Credentials -Authentication Basic -AllowRedirection
        $importResults = Import-PSSession $mSOLSession -AllowClobber
        Write-Verbose $importResults
    }
    Process
    {
        if ($pscmdlet.ShouldProcess('Target', 'Operation'))
        {
        }
    }
    End
    {
    }
}
