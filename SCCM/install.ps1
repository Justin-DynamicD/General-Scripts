#Set Install Variables
$Application = '\setup\nwsapsetup.exe'
$Arguments = '/silent /product="SAPBI"+"SAPGUI710"+"ECL710"+"KW710"+"SAPWUS"+"GUI710ISHMED"+"JNET"+"SAPDTS"+"SCE"'

#If PSVersion is too early, generate $PSScriptRoot variable
If (!$PSScriptRoot) {
	$PSSCriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
    }

#Try the Install and return results
Try {
    $AppFullPath = $PSScriptRoot+'\'+$Application
    start-process -FilePath $AppFullPath -ArgumentList $Arguments -wait
    }
Catch {
    Write-Error -Message $_
    Exit 1
    }
Exit 0
