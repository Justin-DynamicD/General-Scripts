#Import Required modules
If (!(Get-Module ActiveDirectory)) {
    try {import-module ActiveDirectory}
    catch {
        write-error "Unable to load the required Module ActiveDirectory"
        write-verbose $_
        exit 1
        }
    }

Function new-BlizzardResourceGroup {
    Param
        (
	        [Parameter(Mandatory=$true)][String]$GroupName,
            [Parameter(Mandatory=$false)][String]$Path="OU=Groups,OU=Managed Objects,DC=proximus,DC=prosum,DC=com"
        )

    $Test=(Get-ADgroup -filter {SamAccountName -eq $GroupName})
    If ($Test) {
        write-verbose "Group $GroupName already exists"
        }
    Else {
        write-verbose "creating group $GroupName"
        try {
            New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope DomainLocal -Path $Path
            }
        catch {
            write-error "error creating group $GroupName"
            write-verbose $_
            }
        }
    }

function new-BlizzardFranchise {
    Param
    (
	    [Parameter(Mandatory=$true)][String]$Owner="Justin.King_adm",
        [Parameter(Mandatory=$false)][String]$Domain="PROXIMUS.PROSUM.COM",
        [Parameter(Mandatory=$true)][String]$Franchise="WoW",
        [Parameter(Mandatory=$true)][String]$ArtPath="\\es-ssccm-01\DFS",
        [Parameter(Mandatory=$true)][String]$BuildPath="\\es-ssccm-01\DFS",
        [Parameter(Mandatory=$true)][String]$ProjectPath="\\es-ssccm-01\DFS"
    )

    #Validate User Exists
    $Test= Get-ADUser -filter {SamAccountName -eq $Owner}
    If (!($Test)) {
        write-error "User $Owner cannot be found"
        exit 1
        }

    #Validate/Create ResourceGroups
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise-FC
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise-RO
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise-RW
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise.Art-RO
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise.Art-RW
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise.Builds-RO
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise.Builds-RW
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise.Projects-RO
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Franchise.Projects-RW

    #Add Owner to Franchise and grant RW access
    Add-ADGroupMember -Identity RG-FLDR-$Franchise-FC -Members $Owner
    Add-ADGroupMember -Identity RG-FLDR-$Franchise-RW -Members $Owner

    #Create Initial Folder Structure and set permissions
    New-DfsnFolder -Path \\$DOMAIN\Franchise\$Franchise\Art -TargetPath $ArtPath -Description "Folder for Art Library."
    New-DfsnFolder -Path \\$DOMAIN\Franchise\$Franchise\Builds -TargetPath $BuildPath -Description "Folder for Builds."
    New-DfsnFolder -Path \\$DOMAIN\Franchise\$Franchise\Projects -TargetPath $ProjectPath -Description "Folder for Project Stuff."

}

function new-BlizzardDepartment {
    Param
    (
	    [Parameter(Mandatory=$true)][String]$Owner="Justin.King_adm",
        [Parameter(Mandatory=$false)][String]$Domain="PROXIMUS.PROSUM.COM",
        [Parameter(Mandatory=$true)][String]$Department="Cinematics",
        [Parameter(Mandatory=$true)][String]$DepartmentPath="\\es-ssccm-01\DFS"
    )

    #Validate User Exists
    $Test= Get-ADUser -filter {SamAccountName -eq $Owner}
    If (!($Test)) {
        write-error "User $Owner cannot be found"
        exit 1
        }

    #Get Currently Defined Departments
    $DepartmentList = (Get-ChildItem -Path \\ES-SSCCM-01\DSL | ?{ $_.PSIsContainer } | Select-Object Name)

    #Validate/Create ResourceGroups
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Department-FC
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Department-RO
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Department-RW
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Department.Privileged-RO
    new-BlizzardResourceGroup -GroupName RG-FLDR-$Department.Privileged-RW

    #Add Owner to Department and grant RW access
    Add-ADGroupMember -Identity RG-FLDR-$Department-FC -Members $Owner
    Add-ADGroupMember -Identity RG-FLDR-$Department-RW -Members $Owner
    Add-ADGroupMember -Identity RG-FLDR-$Department.Priveleged-RW -Members $Owner

    #Create Initial Folder Structure and set permissions
    New-DfsnFolder -Path \\$DOMAIN\Department\$Department -TargetPath $DepartmentPath -Description "Root Department Path."
    New-Item -ItemType Folder -path "\\$DOMAIN\Department\$Department\Inbox", "\\$DOMAIN\Department\$Department\Priveleged", "\\$DOMAIN\Department\$Department\Collaboration"
    foreach ($AltDept in $DepartmentList) {
        New-Item -ItemType Folder -path "\\$DOMAIN\Department\$Department\Collaboration\$AltDept"
        New-Item -ItemType Folder -path "\\$DOMAIN\Department\$AltDept\Collaboration\$Department"
        }
}

