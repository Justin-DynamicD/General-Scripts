<#
.Synopsis
Allows for consistent Domain Controller deployment and retirement.

.DESCRIPTION
This script will install DFSN and ADFS roles if missing, then promote the server and join the DFSN topology.
The DFSTarget join will return errors while 2003 DCs are present and should be validated after running, 
though it does apear to function properly

This script will skips steps that are already configured and thus is safe to be run multiple times.
Domain Controllers will reboot during configuration, so this can be handy in validating everything is correct.

This script DOES leverage read-host and use a message-box so is not suited for unintended deployment unless 
-unattend flag is used.  If you want full automation don't skip that flag.

.EXAMPLE
InstallDC

This will prompt the user for a safemodepassword then install all missing roles and install the DC.  The computer will reboot when complete.
Run a second time (or if DC has already been promoted), it will add itself to the DFSN targets "BasicNetworkShares" and "FSOutput".

.EXAMPLE
InstallDC -DFSOnly -DFSFolders "FolderA", "FolderB"

This will tell the script to skip domain detection, and attempt to join the DFS targets "FolderA" and "FolderB" only

.EXAMPLE
InstallDC -Uninstall

This will break membership with DFSN as well as demote the server.  SHared direcotries and other contents are left intact.

.PARAMETER SafeModePassword
A Securestring value of the password used for SafeMode recovery of a DC or the new local Administrator password when demoting.

.PARAMETER DFSFolders
An Array of all the DFS root targets in the domain you wish DFSN to join.  This only scans the current domain, as it's not intended 
for use on non-DCs.

.PARAMETER DFSOnly
Switch set to skip any DC checks or changes.  Most logic prvents accidental changes,  so this is more of a paranoia setting

.PARAMETER Unattend
Tells the script to supress prompting for user inputing or generating message boxes. 

.PARAMETER Uninstall
Switch Sets the script to uninstall.  It will not remove roles or local files, only DFSN membership and demote the DC.
#>


#Requires -RunasAdministrator

param(
    [parameter(Mandatory=$false)]
    [System.Security.SecureString]$SafeModePassword,

    [parameter(Mandatory=$false)]
    [array]$DFSFolders = @("BasicNetworkShares","FSOutput"),

    [parameter(Mandatory=$false)]
    [switch]$DFSOnly,

    [parameter(Mandatory=$false)]
    [switch]$Unattend,

    [parameter(Mandatory=$false)]
    [switch]$Uninstall
    )

#Grab Environment Varibales
$IsDomainController = If ((Get-WmiObject win32_computersystem).domainrole -gt "3") {$true} else {$false}
$DomainName = (Get-ADDomain).NetBIOSName
$DomainFQDN = (Get-ADDomain).DNSRoot

#Verify $SafeModePassword is present if uninstall or $DFSOnly isn't set
If ((!$DFSOnly -or ($Uninstall -and $IsDomainController) -or (!$Uninstall -and !$IsDomainController)) -and !$SafeModePassword) {
    If (!$Unattend) {
        [System.Security.SecureString]$SafeModePassword = Read-Host -Prompt "Local/SafeMode Administrator Password" -AsSecureString
        }
    Else {
        write-error "-SafeModePassword is required with this combination of parameters" -ErrorAction Stop
        }
    }

#Validate roles are installed
If (!(Get-WindowsFeature AD-Domain-Services | where-object {$_.InstallState -eq "Installed"}) -and !$Uninstall -and !$DFSOnly) {
    try {Write-Verbose "Domain services missing, installing..."
        Add-WindowsFeature AD-Domain-Services -IncludeManagementTools
        }
    catch {
        Write-Error "Cannot add domain services, stopping" -ErrorAction Stop
        }
    }

If (!(Get-WindowsFeature FS-DFS-Namespace | where-object {$_.InstallState -eq "Installed"}) -and !$Uninstall) {
    try {Write-Verbose "DFS services missing, installing..."
        Add-WindowsFeature FS-DFS-Namespace -IncludeManagementTools
        }
    catch {
        Write-Error "Cannot add dfs services, stopping" -ErrorAction Stop
        }
    }

#Create required shares for DFS namespace
write-verbose "creating DFS_shares if missing"
If (!(test-path $env:SystemDrive\DFS_Shares) -and !$Uninstall) {new-item -ItemType Directory $env:SystemDrive\DFS_Shares -force | Out-Null}

#loop array and create all needed subfolders
If (!$Uninstall) {
    foreach ($DFSName in $DFSFolders) {
        If (!(test-path $env:SystemDrive\DFS_Shares\$DFSName)) {
            write-verbose "creating $DFSName..."
            new-item -ItemType Directory $env:SystemDrive\DFS_Shares\$DFSName -force | Out-Null
            $ACL= (Get-Item $env:SystemDrive\DFS_Shares\$DFSName).GetAccessControl('Access')
            $AR = New-Object System.Security.AccessControl.FileSystemAccessRule("$DomainName\Domain Users", 'ReadandExecute', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            $ACL.SetAccessRule($AR)
            Set-Acl -path $env:SystemDrive\DFS_Shares\$DFSName -AclObject $ACL
            New-SmbShare -Name $DFSName -Path $env:SystemDrive\DFS_Shares\$DFSName -ChangeAccess Everyone | Out-Null
            }#End If Loop
        }#End ForEach Loop
    }#End Uninstall Condition

#Check if AD has been configured, install if not.
If (!$IsDomainController -and !$Uninstall -and !$DFSOnly) {
    If (!$Unattend) {[System.Windows.Forms.MessageBox]::Show("The Domain Controller role is about to be installed.  This will cause your server to reboot when replication completes.  Please re-run this script to finish installation of servies that depend on the DC.", "Status")}
    Else {Write-Verbose "System is being promoted to a DC and will reboot when complete..."}
    Install-ADDSDomainController -InstallDns -DomainName $DomainFQDN -SafeModeAdministratorPassword $SafeModePassword
    break
    }

#Join/Break each DFS Target 
If (!$Uninstall) {
    foreach ($DFSName in $DFSFolders) {
        if (!((Get-DfsnRootTarget -Path "\\$DomainName\$DFSName").TargetPath | where { $_ -like "\\$env:COMPUTERNAME*"})) {
            write-verbose "Attempting to join \\$DomainName\$DFSName"
            New-DfsnRootTarget -path "\\$DomainName\$DFSName" -TargetPath "\\$env:COMPUTERNAME\$DFSName"
            }#End If Loop
        }#End ForEachLoop
    }#End Uninstall exception
Else {
    foreach ($DFSName in $DFSFolders) {
        if ((Get-DfsnRootTarget -Path "\\$DomainName\$DFSName").TargetPath | where { $_ -like "\\$env:COMPUTERNAME*"}) {
            write-verbose "Attempting to remove \\$DomainName\$DFSName"
            Remove-DfsnRootTarget -path "\\$DomainName\$DFSName" -TargetPath "\\$env:COMPUTERNAME\$DFSName"
            }#End If Loop
        }#End ForEachLoop
    }

#If Uninstall was set, demote the domain controlker
If ($IsDomainController -and $Uninstall -and !$DFSOnly) {
    If (!$Unattend) {[System.Windows.Forms.MessageBox]::Show("The Domain Controller role is about to be demoted.  This will cause your server to reboot when replication completes.", "Status")}
    Else {write-verbose "Server is being demoted, a reboot will occur shortly"}
    Uninstall-ADDSDomainController -LocalAdministratorPassword $SafeModePassword
    break
    }
