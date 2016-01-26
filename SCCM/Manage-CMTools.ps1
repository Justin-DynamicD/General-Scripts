    param (
         [Parameter(Mandatory=$true)][ValidateSet("CMConsole","MDT")][System.String]$Application,
         [Parameter(Mandatory=$false)][System.String]$SiteServer,
         [Parameter(Mandatory=$false)][System.String]$DeploymentShare,
         [Parameter(Mandatory=$false)][ValidateSet("Install","Remove","Validate")][System.String]$Mode="Install"
        )

    #If PSVersion is too early, generate $PSScriptRoot variable
    If (!$PSScriptRoot) {
        $PSSCriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
        }
    
    #Set variables to starting values
    Set-Location -Path $PSScriptRoot
    $CMConsoleInstaller='ConsoleSetup.exe'
    $MDTInstaller='MicrosoftDeploymentToolkit2013_x64.msi'
    $mdtpresent=$false
    $cmpresent=$false
    $mergetools=$false

    #Close early if the right combination of parameters arn't present
    If (($Application -eq "CMConsole") -and ($Mode -eq "Install") -and (!$SiteServer)) {
        write-error "-SiteServer is required in order to perform a silent install"
        exit 1
        }
    If (($Application -eq "MDT") -and ($Mode -eq "Install") -and (!$DeploymentShare)) {
        write-error "-DeploymentShare is required in order to perform a silent install"
        exit 1
        }

    #Get all paths and determine if apps are present
    $mdtPath = (get-itemproperty "hklm:\Software\Microsoft\Deployment 4" -name Install_Dir -ErrorAction SilentlyContinue).Install_Dir
    if ($mdtPath -eq $null) {
        write-verbose "Unable to locate the Microsoft Deployment Toolkit installation directory."
        }
    else {
        $mdtpresent=$true
        }
    $cm07Path = (get-itemproperty "hklm:\Software\Microsoft\ConfigMgr\Setup" -name "UI Installation Directory" -ErrorAction SilentlyContinue).'UI Installation Directory'
    if ($cm07Path -ne $null) {
        Write-verbose "Found CM07 at $cm07Path"
        }
    else {
        $cm07Path = (get-itemproperty "hklm:\Software\wow6432node\Microsoft\ConfigMgr\Setup" -name "UI Installation Directory" -ErrorAction SilentlyContinue).'UI Installation Directory'
        if ($cm07Path -ne $null) {
            Write-verbose "Found CM07 (32-bit) at $cm07Path"
            }
        }
    $cm12Path = (get-itemproperty "hklm:\Software\Microsoft\ConfigMgr10\Setup" -name "UI Installation Directory" -ErrorAction SilentlyContinue).'UI Installation Directory'
    if ($cm12Path -ne $null) {
        Write-Verbose "Found CM12 at $cm12Path"
        }
    else {
        $cm12Path = (get-itemproperty "hklm:\Software\wow6432node\Microsoft\ConfigMgr10\Setup" -name "UI Installation Directory" -ErrorAction SilentlyContinue).'UI Installation Directory'
        if ($cm12path -ne $null) {
            Write-Verbose "Found CM12 (32-bit) at $cm12Path"
            }
        }
    if (!$cm12Path -and !$cm07path) {
        write-verbose "Unable to locate any versions of SCCM Console"
        }
    else {
        $cmpresent=$true
        }

    #Report Back if Present and Mode is set to Validate
    If ($Mode -eq "Validate") {
        switch ($Application) {
            CMConsole {
                Return $cmpresent
                }
            MDT {
                Return $mdtpresent
                }
            }
        }

    #Install Application if mode is set to Install
    If ($Mode -eq "Install") {
        If (($Application -eq "CMConsole") -and ($cmpresent -eq $true)) {
            Write-verbose "CMConsole is already installed"
            exit 0
            }
        If (($Application -eq "MDT") -and ($mdtpresent -eq $true)) {
            write-verbose "MDT is already installed"
            exit 0
            }
        switch ($Application) {
            CMConsole {
                try {
                    $Arguments = 'EnableSQM=0 TARGETDIR="' + $env:programfiles + '\ConfigmanConsole" DEFAULTSiteServerName=' + $SiteServer + ' /q'
                    start-process -FilePath $CMConsoleInstaller -ArgumentList $Arguments -wait
                    $cm12Path = (get-itemproperty "hklm:\Software\wow6432node\Microsoft\ConfigMgr10\Setup" -name "UI Installation Directory" -ErrorAction SilentlyContinue).'UI Installation Directory'
                    }
                catch {
                    Write-Error -Message $_                    Write-Error "Error installing CMConsole"
                    exit 1
                    }
                if ($mdtpresent -eq $true) {
                   $mergetools = $true 
                    }
                }
            MDT {
                try {
                    $Arguments = '/q REBOOT=ReallySuppress /I '+ $MDTInstaller
                    start-process -FilePath "msiexec.exe" -ArgumentList $Arguments -wait
                    $mdtPath = (get-itemproperty "hklm:\Software\Microsoft\Deployment 4" -name Install_Dir -ErrorAction SilentlyContinue).Install_Dir
                    }
                catch {
                    Write-Error -Message $_                    Write-Error "Error installing MDT"
                    exit 1
                    }
                Import-Module ($mdtPath + "\bin\MicrosoftDeploymentToolkit.psd1")
                new-PSDrive -Name "DS001" -PSProvider "MDTProvider" -Root $DeploymentShare -Description "MDT Deployment Share" | add-MDTPersistentDrive
                if ($cmpresent -eq $true) {
                   $mergetools = $true 
                    }
                }
            }
        }

    #If MergeTools is set to True and mode is set to Install, integrate MDT with CMConsole
    If (($Mode -eq "Install") -and ($mergetools -eq $true)) {
        if ($cm07Path -ne $null) {
            Write-Verbose "Integrating MDT into the ConfigMgr 2007 console"
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.SCCMActions.dll" -Destination "$cm07Path\Bin\Microsoft.BDD.SCCMActions.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.Workbench.dll" -Destination "$cm07Path\Bin\Microsoft.BDD.Workbench.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.Wizards.dll" -Destination "$cm07Path\Bin\Microsoft.BDD.Wizards.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.PSSnapIn.dll" -Destination "$cm07Path\Bin\Microsoft.BDD.PSSnapIn.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.Core.dll" -Destination "$cm07Path\Bin\Microsoft.BDD.Core.dll" -Force
            Copy-Item -Path "$mdtPath\Templates\Extensions" -Destination "$cm07Path\XmlStorage\Extensions" -Force -Recurse
            }

        if ($cm12Path -ne $null) {
            Write-Verbose "Integrating MDT into the ConfigMgr 2012 console"
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.CM12Actions.dll" -Destination "$cm12Path\Bin\Microsoft.BDD.CM12Actions.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.Workbench.dll" -Destination "$cm12Path\Bin\Microsoft.BDD.Workbench.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.ConfigManager.dll" -Destination "$cm12Path\Bin\Microsoft.BDD.ConfigManager.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.CM12Wizards.dll" -Destination "$cm12Path\Bin\Microsoft.BDD.CM12Wizards.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.PSSnapIn.dll" -Destination "$cm12Path\Bin\Microsoft.BDD.PSSnapIn.dll" -Force
            Copy-Item -Path "$mdtPath\Bin\Microsoft.BDD.Core.dll" -Destination "$cm12Path\Bin\Microsoft.BDD.Core.dll" -Force
            Copy-Item -Path "$mdtPath\Templates\CM12Extensions" -Destination "$cm12Path\XmlStorage\Extensions" -Force -Recurse
            }

        }

    #If both tools are present and mode is set to Remove, break integration between MDT and CMConsole first
    If (($mdtpresent -eq $true) -and ($cmpresent -eq $true) -and ($Mode -eq "Remove")) {
        if ($cm07Path -ne $null) {
            Write-Verbose "Removing MDT extensions from the ConfigMgr 2007 console"
            Remove-Item -Path "$cm07Path\Bin\Microsoft.BDD.SCCMActions.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm07Path\Bin\Microsoft.BDD.Workbench.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm07Path\Bin\Microsoft.BDD.Wizards.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm07Path\Bin\Microsoft.BDD.PSSnapIn.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm07Path\Bin\Microsoft.BDD.Core.dll" -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path "$mdtPath\Templates\CM12Extensions" -Recurse | foreach {Get-ChildItem $cm07Path\xmlstorage\extensions\$_ -Recurse  -ErrorAction SilentlyContinue | Remove-Item -force -ErrorAction SilentlyContinue}
            }

        if ($cm12Path -ne $null) {
            Write-Verbose "Removing MDT extensions from the ConfigMgr 2007 console"
            Remove-Item -Path "$cm12Path\Bin\Microsoft.BDD.SCCMActions.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm12Path\Bin\Microsoft.BDD.Workbench.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm12Path\Bin\Microsoft.BDD.Wizards.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm12Path\Bin\Microsoft.BDD.PSSnapIn.dll" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$cm12Path\Bin\Microsoft.BDD.Core.dll" -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path "$mdtPath\Templates\CM12Extensions" -Recurse | foreach {Get-ChildItem $cm12Path\xmlstorage\extensions\$_ -Recurse -ErrorAction SilentlyContinue | Remove-Item -force -ErrorAction SilentlyContinue}
            }
        }

    #Remove selected Application
    If ($Mode -eq "Remove") {
        If (($Application -eq "CMConsole") -and ($cmpresent -eq $false)) {
            write-verbose "CMConsole was not found"
            exit 0
            }
        If (($Application -eq "MDT") -and ($mdtpresent -eq $false)) {
            write-verbose "MDT was not found"
            exit 0
            }

    switch ($Application) {
            CMConsole {
                try {
                    $Arguments = '/uninstall /q'
                    start-process -FilePath $CMConsoleInstaller -ArgumentList $Arguments -wait
                    $cm12Path = $null
                    }
                catch {
                    Write-Error -Message $_                    Write-Error "Error uninstalling CMConsole"
                    Exit 1
                    }
                }
            MDT {
                try {
                    $Arguments = '/x {CFF8B5ED-0A4D-4EDD-9159-32FE1D31C9E3} /q'
                    start-process -FilePath "msiexec.exe" -ArgumentList $Arguments -wait
                    $mdtPath = $null
                    }
                catch {
                    Write-Error -Message $_                    Write-Error "Error uninstalling MDT"
                    Exit 1
                    }
                }

        }
    }

    Exit 0
