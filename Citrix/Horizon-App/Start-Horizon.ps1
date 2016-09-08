#These variables should be updated as needed for the environment 
$Horizon = "legaclt.exe" #name of the executable
$HorizonPath = "C:\Users\Public\Rocket Software\LegaSuite Windows Client\7.2456.1.34" #Path to the executable
$HorizonArguements = '/h10.118.32.30 /GWorD="'+$HorizonPath+'\Production" /M2924 /GPanR=1 /GFonR=1 /GDisPR=1 /GCCSID=37 /GShowEmulator=1 /GMAIWS=59 /Gdevns=off' #startup Flags
$CSVFile = "UserName-Mapping.csv" #name of the CSV File
$CSVFilePath = "\\prfile01\Install\Horizon-Run" #Path to the CSV
[switch]$CreateCSVIfMissing = $false #Is the script is allowed to create an empty CSV sample file if one is missing
$WSTemplate = "WS-Template.ws" #name of the template used to generate the WS launch file
$WSTemplatePath = "\\prfile01\Install\Horizon-Run" #Path to the WS launch file
$TempPath = $env:TEMP #Temp directory the launch file will be placed in
[switch]$DoNotCleanupTemp = $false #Is the script going to remove any files it creates

#Import the CSV File and store variables in $csvdata
$FileName= $CSVFilePath+"\"+$CSVFile
If(!(Test-Path -Path ($FileName)) -and ($CreateCSVIfMissing)) {
    try {
        New-Item ($FileName) -type file -force -ErrorAction "STOP" | Out-Null
        $NewLine = "UserName,DJNumber,PNumber"
        $NewLine | add-content -path $FileName -ErrorAction "STOP"
        }
    catch {
            $E = $_.Exception.GetBaseException()
            write-error $E.ErrorInformation.Description -ErrorAction "Stop"
            }
    } #End If
elseif (!(Test-Path -Path ($FileName))) {
    Write-Error "$FileName not found.  Please update the script." -ErrorAction "Stop"
    }

Try {
    $CSVData = (import-csv $FileName | where-object {$_.UserName -eq $env:username})[0]
    }
Catch {
    Write-Error "Cannot find any entries mathcing workstation $env:clientname" -ErrorAction "Stop"
    }

#Rebuild the Horizon launch Arguements adding the DJNumber then launch the program
$FileName = $HorizonPath+"\"+$Horizon
$HorizonArguements = $HorizonArguements+" /Gdevn="+$CSVData.DJNumber
try {
    Start-Process -FilePath $FileName -ArgumentList $HorizonArguements -ErrorAction "Stop"
    }
catch {
     $E = $_.Exception.GetBaseException()
     write-error $E.ErrorInformation.Description -ErrorAction "Stop"   
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
        write-verbose $WSFileName
        (Get-Content $SourceName).replace('[DJNumber]',$CSVData.DJNumber).replace('[PNumber]',$CSVData.PNumber) | Set-Content $WSFileName -force -ErrorAction "Stop"
        Start-Process -FilePath $WSFileName -ErrorAction "Stop"
        }
    catch {
        $E = $_.Exception.GetBaseException()
        write-error $E.ErrorInformation.Description -ErrorAction "Stop" 
        }
    } #End Take PNumber Action

#Cleanup the WSFileName file if it exists
If ($WSFileName -and !($DoNotCleanupTemp)) {
    Start-Sleep -s 3 #Do not remove the file for 3 seconds to allow the program to launch
    Remove-Item -Path $WSFileName -force
    }
