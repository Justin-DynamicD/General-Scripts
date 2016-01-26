Param
(
	[Parameter(Mandatory=$false)][Alias('Dir')][Array]$DirList="\\proximus.prosum.com\DFS",
	[Parameter(Mandatory=$false)][Alias('OutFile')][String]$Output = "Results.csv",
    [Parameter(Mandatory=$false)][Alias('OutType')][ValidateSet("CSV","HTML")][String]$OutputType = "CSV",
    [Parameter(Mandatory=$false)][String]$Depth="10"
)

#Check for and load NTFSSecurity module if it's missing
IF (!(Get-Module NTFSSecurity)) {
    write-verbose "Importing the NTFSSecurity Module ..."
    Try {
        Import-Module .\NTFSSecurity\NTFSSecurity.psd1
        }
    Catch {
        Throw "unable to load the NTFSSecurity PowerShell Module. Please ensure the modules exist under .\NTFSSecurity or import them manually"
        }
    }

IF (!(Get-Module DFSN)) {
    write-verbose "Importing the DFSN Module ..."
    Try {
        Import-Module DFSN
        }
    Catch {
        Throw "unable to load the DFSN PowerShell Module.  Please use add-remove programs to add DFS name space management tools or run from a server with these tools installed"
        }
    }

#This function solves the lack of depth selection in Get-ChildItem (won't be needed in PSv5)
write-verbose "creating custom function to crawl folders to specified depth ..."
function Get-ChildItemToDepth {
  param(
    [String]$Path = $PWD,
    [Byte]$ToDepth = 255,
    [Byte]$CurrentDepth = 0,
    [Switch]$DebugMode
  )
 
  $CurrentDepth++
  if ($DebugMode) { $DebugPreference = "Continue" }
 
  Get-ChildItem $Path -Directory | ForEach-Object {
    $_.FullName
    if ($_.PsIsContainer) {
      if ($CurrentDepth -le $ToDepth) {
        # Callback to this function
        Get-ChildItemToDepth -Path $_.FullName -ToDepth $ToDepth -CurrentDepth $CurrentDepth
      } else {
        Write-Debug $("Skipping GCI for Folder: $($_.FullName) " +
          "(Why: Current depth $CurrentDepth vs limit depth $ToDepth)")
      }
    }
  }
}

#Start with an empty array to be populated
[array]$ChildList = $null

#Add each defined URL to array with child paths
write-verbose "Building full directory tree (this could take awhile) ..."

foreach($DL in $DirList) {
    $ChildList+=$DL
    $ChildList+=(Get-ChildItemToDepth -Path $DL -ToDepth $Depth)
    }

#Because variables will be added to, start clean
[array]$DFSConvert = @()
[array]$FinalList=$null
$i=0

#Begin a ForEach loop to update childlist to finallist...
write-verbose "Gathering Information..."
$ChildList | ForEach-Object {
    $Source = $_
    $newpath=$null
    $newvalue=$null

    #Check path for DFS Namespace and update if present
    If (Get-DfsnFolderTarget $Source -ErrorAction SilentlyContinue| where {$_.State -eq "Online"}) {
        $NewPath = (Get-DfsnFolderTarget $Source -ErrorAction SilentlyContinue| where {$_.State -eq "Online"})[0]
        write-verbose "path needs to be updated to $NewPath ..."
        $DFSConvert+=$NewPath
        $newvalue = $newpath.TargetPath
        }

    #If current depth is not a namespace, check if full path contains a namespace and adjust UNC
    ElseIf ($DFSConvert)  {
        ForEach ($DFS in $DFSConvert) {
            $Test=$DFS.Path
            $Testinj = [regex]::escape($Test)
            $TargetTest = $DFS.TargetPath
            If ($Source -like $Test+"*") {
                Write-verbose "need to update $_ with $targettest..."
                $NewValue = ($_ -replace($Testinj,$TargetTest))
                }
            }
        }

    #if after both checks there's no match, set the new value to equal original
    if ($newvalue -eq $null) {$newvalue=$Source}

    #with proper UNC known, check permisssions then build custom object to store all results
    write-verbose "checking for explicit permissions on folder..."
    $perms = (Get-NTFSAccess -Path $newvalue | where {$_.IsInherited -eq $false} | Select Account, AccessRights)
    write-verbose "Saving information ..."
    $tempobject = New-Object -TypeName PSObject
    $tempobject | Add-Member -MemberType NoteProperty -Name NameSpace -Value $ChildList[$i]
    $tempobject | Add-Member -MemberType NoteProperty -Name UNC -Value $NewValue
    $tempobject | Add-Member -MemberType NoteProperty -Name Permissions -Value $perms
    [array]$FinalList+=$tempobject

    #incriment counter for next line in loop
    $i++
    }

#Gather all non-inherited permissions and save based on defined type
write-verbose "Generating report..."
Switch ($OutputType) {
    HTML {$finallist | select NameSpace, UNC, @{Label='Accounts';EXPRESSION={($_.permissions.account.accountname | out-string).Trim()}}, @{Label='Permissions';EXPRESSION={($_.permissions.AccessRights | out-string).Trim()}} | ConvertTo-Html | Out-File $Output}
    CSV {$finallist | select NameSpace, UNC, @{Label='Accounts';EXPRESSION={($_.permissions.account.accountname | out-string).Trim()}}, @{Label='Permissions';EXPRESSION={($_.permissions.AccessRights | out-string).Trim()}} | Export-Csv $Output}
    }

write-verbose "Done!"
