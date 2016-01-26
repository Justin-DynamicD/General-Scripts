Param
(
	[Parameter(Mandatory=$true)][Alias('TDir')][String]$TargetDir,
	[Parameter(Mandatory=$false)][Alias('Log')][String]$LogFile = "update-results.log",
    [Parameter(Mandatory=$false)][String]$OU,
    [Parameter(Mandatory=$false)][String]$User,
    [Parameter(Mandatory=$false)][Alias('SDir')][String]$SourceDir,
    [Parameter(Mandatory=$false)][Alias('File')][String]$TxtFile,
    [Parameter(Mandatory=$false)][Switch]$NoCopy = $false,
    [Parameter(Mandatory=$false)][Switch]$NoAD = $false
)

#Check for and ensure only one source is selected, not multiple
[int]$SourceCount=0
IF ($OU) {$SourceCount=$SourceCount+1}
IF ($User) {$SourceCount=$SourceCount+1}
IF ($SourceDir) {$SourceCount=$SourceCount+1}
IF ($TxtFile) {$SourceCount=$SourceCount+1}

IF (!($SourceCount -eq 1)){
    Write-Error "$SourceCount sources were provided, please try again and use -OU, -User, -SourceDir, or -TxtFile to designate a single source of data"
    exit
    } #End Count Error

#Validate Paths or files are "real"
IF (($SourceDir) -and !(Test-Path -path $SourceDir)) {
    Write-Error "$SourceDir cannot be found.  Plesae check the path and try again"
    exit
    }
IF (!(Test-Path -path $TargetDir)) {
    Write-Error "$TargetDir cannot be found.  Plesae check the path and try again"
    exit
    }
IF (!(Test-Path $TxtFile)) {
    Write-Error "$TxtFile cannot be found.  Plesae check the path and try again"
    exit
    }

#Go through each type of search criteria and build an array of user objects
[array]$UserList=$Null

IF ($User) {
    Try {
        $UserList=(Get-ADUser $User)
        }
    Catch {
        Write-Error "Cannot find the account $User.  Exiting."
        }
    } #End User Variable

IF ($OU) {
    Try {
        $UserList=(Get-ADUser -Filter * -SearchBase $OU)
        }
    Catch {
        Write-Error "Cannot find any accounts under $OU.  Exiting."
        }
    } #End OU Variable

IF ($SourceDir -or $TxtFile) {
    IF ($SourceDir) {$EnumerateDir=(Get-ChildItem $SourceDir -Directory).Name}
    IF ($TxtFile) {$EnumerateDir=Get-Content $TxtFile}
    ForEach ($UserCheck in $EnumerateDir) {        
        Try {
            $UserList+=(Get-ADUser $UserCheck)
            }
        Catch {
            Write-Warning "Cannot map $UserCheck to a user."
            }
        } #Finish Validating dir to users
    IF ($UserList -eq $Null) {
        Write-Error "No Users were matched!"
        Exit
        } #End Empty List Error
    } #End SourceDir Variable

#This Section "Does the Work" by going through the user object list and applying changes.
foreach($Account in $UserList) {
    
    #Set simple variables
    $DN=$Account.DistinguishedName
    $Set_TSProfile = [ADSI] "LDAP://$DN"
    $TsProfilePath=$TargetDir + $Account.SamAccountName
    
    #Attempt to gather original path which may error
    Try {
        $OldTsProfilePath=$Set_TsProfile.psbase.invokeGet("TerminalServicesProfilePath")
        }
    Catch {
        Write-Verbose "Unable to find and old path, setting to net-new"
        $OldTsProfilePath="Net-New"
        }

    #Update User Object unless "noAD" switch is set
    IF ($NoAD -eq $false) {
        Write-Verbose "updating User $DN"
        $Set_TsProfile.psbase.invokeSet("TerminalServicesProfilePath",$TsProfilePath)
        $Set_TsProfile.setinfo()
        } #Update AD

    #Create New Path if it's missing
    IF (!(Test-Path -path $TsProfilePath)) {
        New-Item $TsProfilePath -ItemType Directory -force
        } #End Create Path

    #Update file location unless "NoCopy" switch is set
    IF (($NoCopy -eq $false) -and !($OldTsProfilePath -eq "Net-New")) {
        Write-verbose "attempting to copy data from $OldTsProfilePath"
        Copy-Item -Path $OldTsProfilePath -Destination $TsProfilePath -Recurse -Force
        } #Update Files

    } #End ForEach user in list 