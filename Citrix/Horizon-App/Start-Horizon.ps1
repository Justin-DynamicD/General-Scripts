#These variables should be updated as needed for the environment 
$Horizon = "legaclt.exe"
$HorizonPath = "C:\Users\Public\Rocket Software\LegaSuite Windows Client\7.2456.1.34"
$HorizonArguements = '/h10.118.32.30 /GWorD="'+$HorizonPath+'\Production" /M2924 /GPanR=1 /GFonR=1 /GDisPR=1 /GCCSID=37 /GShowEmulator=1 /GMAIWS=59 /Gdevns=off'
$CSVFile = "Workstation-Mapping.csv"
$CSVFilePath = $PSScriptRoot
[switch]$CreateCSVIfMissing = $false
$WSTemplate = "WS-Template.ws"
$WSTemplatePath = $PSScriptRoot
$TempPath = $env:TEMP
[switch]$DoNotCleanupTemp = $false

#Import the CSV File and store variables in $csvdata
$FileName= $CSVFilePath+"\"+$CSVFile
If(!(Test-Path -Path ($FileName)) -and ($CreateCSVIfMissing)) {
    try {
        New-Item ($FileName) -type file -force -ErrorAction STOP | Out-Null
        $NewLine = "WorkstationName,DJNumber,PNumber"
        $NewLine | add-content -path $FileName -ErrorAction STOP
        }
    catch {
            $E = $_.Exception.GetBaseException()
            write-error $E.ErrorInformation.Description -ErrorAction Stop
            }
    } #End If
elseif (!(Test-Path -Path ($FileName))) {
    Write-Error "$FileName not found.  Please update the script." -ErrorAction Stop
    }
$CSVData = (import-csv $FileName | where-object {$_.WorkstationName -eq $env:computername})[0]

#Rebuild the Horizon launch Arguements adding the DJNumber then launch the program
$FileName = $HorizonPath+"\"+$Horizon
$HorizonArguements = $HorizonArguements+" /Gdevn="+$CSVData.DJNumber
try {
    Start-Process -FilePath $FileName -ArgumentList $HorizonArguements -ErrorAction Stop
    }
catch {
     $E = $_.Exception.GetBaseException()
     write-error $E.ErrorInformation.Description -ErrorAction Stop   
    }

#If the PNumber is defined, create the ws template and launch
If ($CSVData.PNumber) {
    try {
        #Generate the random filename
        $Rset=$Null
        $WSFileName=$Null
        $set = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray()
        for ($i=1; $i -le 6; $i++) {
            $Rset += $set | Get-Random
        }
        $WSFileName = $TempPath+"\"+$Rset + ".ws"
        $SourceName = $WSTemplatePath+"\"+$WSTemplate
        (Get-Content $SourceName).replace('[DJNumber]',$CSVData.DJNumber).replace('[PNumber]',$CSVData.PNumber) | Set-Content $WSFileName -force -ErrorAction Stop
        Start-Process -FilePath $WSFileName -ErrorAction Stop
        }
    catch {
        $E = $_.Exception.GetBaseException()
        write-error $E.ErrorInformation.Description -ErrorAction Stop 
        }
    } #End Take PNumber Action

#Cleanup the WSFileName file if it exists
If ($WSFileName -and !($DoNotCleanupTemp)) {
    Remove-Item -Path $WSFileName -force
    }
