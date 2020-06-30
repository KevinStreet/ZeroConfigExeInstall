<#
.SYNOPSIS
	This script is used to determine which installer an executable is using and silently install or uninstall it using default parameters.
	# LICENSE #
	Zero-Config Executable Installation. 
	Copyright (C) 2020 - Kevin Street.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. 
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
    This script attempts to find which installer technology an software product is using and then installdes it silently using the default parameters.
    The script supports the following installer products:
        • Windows Installer
        • NSIS (Nullsoft Scriptable Install System)
        • Inno Setup
        • InstallShield
        • WiX Burn
        • Wise
        • InstallAWARE
        • install4j
        • Setup Factory

    It also has special logic for installing the following products:
        • Microsoft Office 365 click-to-run
        • Microsoft Windows update files (.msu)

    This script is designed to be used as an extension of the Powershell App Deployment Toolkit.
.PARAMETER deploymentType
	Install or uninstall an application.
    This paramter is passed to the script when invoked from the PSADT.
.EXAMPLE
    ZeroConfigExeInstallation.ps1 -deploymentType Install
.EXAMPLE
    ZeroConfigExeInstallation.ps1 -deploymentType Uninstall
.NOTES
    Script version: 1.1.0
    Release date: 30/06/2020.
    Author: Kevin Street.
.LINK
	https://kevinstreet.co.uk

#Requires -Version 2.0
#>

[CmdletBinding()]

param (
    [parameter(Mandatory=$false)]
    [String[]]$TestSupportedInstallerTypePath
)

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
##region VariableDeclaration

## Figure out whether this script has been invoked by the Powershell App Deployment Toolkit.
If ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $appDeployToolkitExtName
}
elseif ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
	Write-Host "This script is designed to be used as an extension of the Powershell App Deployment Toolkit.`nPlease visit https://kevinstreet.co.uk/zero-config-executable-installation or https://github.com/KevinStreet/ZeroConfigExeInstall for details on how to integrate it. `n`nIf you intended to test whether your application installer is supported by this script, please run it again using the following arguments: `nZeroConfigExeInstallation.ps1 -TestSupportedInstallerTypePath 'PathToInstallerExe'"
    Exit
}

## Do not declare variables if $TestSupportedInstallerTypePath is set as that means the user is using this script to test compatibility of a 
## particular installer executable.
if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
    [string]$appDeployToolkitExtName = 'ZeroConfigExe'
    [string]$appDeployExtScriptFriendlyName = 'Zero-Config Executable Installation'
    [version]$appDeployExtScriptVersion = [version]'1.1.0'
    [string]$appDeployExtScriptDate = '30/06/2020'

    ## Check for Exe installer and modify the installer path accordingly.
    ## If multiple .exe files are found, the user may be including both x86 and x64 installers. Check for "86" or "32" and "64" in the names.
    ## If multiple .exe files but they are not for x86 and x64, then look for setup.exe or install.exe and use those.
    ## If neither exist the user must specify the installer executable in the $installerExecutable variable in Deploy-Application.ps1.
    if ([string]::IsNullOrEmpty($installerExecutable)) {
        [array]$exesInPath = @((Get-ChildItem -Path "$dirFiles\*.exe") | Select-Object -Expand Name)
        if ($exesInPath.Count -gt 1) {
            if ((($exesInPath -like "*86*") -or $exesInPath -like "*32*") -and ($exesInPath -like "*64*")) {
                if ($Is64Bit) {
                    [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -like '*64*') } | Select-Object -ExpandProperty 'FullName'
                    Write-Log -Message "x86 and x64 installers found. This system is x64, so x64 installer will be used." -Source $appDeployToolkitExtName
                }

                else {
                    [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and (([IO.Path]::GetFileName($_.Name) -like '*86*') -or ([IO.Path]::GetFileName($_.Name) -like '*32*')) } | Select-Object -ExpandProperty 'FullName'
                    Write-Log -Message "x86 and x64 installers found. This system is x86, so x86 installer will be used." -Source $appDeployToolkitExtName
                }
            }

            elseif ($exesInPath -contains "setup.exe") {
                [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'setup.exe') } | Select-Object -ExpandProperty 'FullName'
            }

            elseif ($exesInPath -contains "install.exe") {
                [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'install.exe') } | Select-Object -ExpandProperty 'FullName'
            }

            else {
                Write-Log -Message "Multiple .exe files found but not sure which to use as the installer. The installer .exe must be specified in the $installerExecutable variable in Deploy-Application.ps1." -Source $appDeployToolkitExtName
            }
        }

        else {
            [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetExtension($_.Name) -eq '.exe') } | Select-Object -ExpandProperty 'FullName' -First 1
        }
    }

    else {
    
        ## If the user manually specified which executable file to use, but did not add the extension (.exe), add it here so it still works.
        if (-not ($installerExecutable -like "*.exe")) {
            $installerExecutable = $installerExecutable + '.exe'
        }

        [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq "$installerExecutable") } | Select-Object -ExpandProperty 'FullName'
    }

    if (-not ([string]::IsNullOrEmpty($defaultExeFile))) {
        Write-Log -Message "Installer executable found: $defaultExeFile." -Source $appDeployToolkitExtName
    }

    else {
        Write-Log -Message "No installer executable was found." -Source $appDeployToolkitExtName
    }

    ## Check for Msu installer and modify the installer path accordingly.
    ## If multiple .msu files are found inform the user they must specify which one they want in $installerExecutable in Deploy-Application.ps1.
    if ([string]::IsNullOrEmpty($defaultExeFile)) {
        if ([string]::IsNullOrEmpty($installerExecutable)) {
            [array]$msusInPath = @((Get-ChildItem -Path "$dirFiles\*.msu") | Select-Object -Expand Name)
            if ($msusInPath -gt 1) {
                Write-Log -Message "Multiple .msu files found but not sure which one to use. Please reduce to one .msu or specify which .msu to use in the $installerExecutable variable in Deploy-Application.ps1." -Source $appDeployToolkitExtName
            }

            else {
                [string]$defaultMsuFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetExtension($_.Name) -eq '.msu') } | Select-Object -ExpandProperty 'FullName' -First 1
            }
        }

        else {

            ## If the user manually specified which Msu installer file to use, but did not add the extension (.msu), add it here so it still works.
            if (-not ($installerExecutable -like "*.msu")) {
                $installerExecutable = $installerExecutable + '.msu'
            }

            [string]$defaultMsuFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq "$installerExecutable") } | Select-Object -ExpandProperty 'FullName'
        }         
    }

    if (-not ([string]::IsNullOrEmpty($defaultMsuFile))) {
        Write-Log -Message "Microsoft update installer found: $defaultMsuFile." -Source $appDeployToolkitExtName
    }
}

##endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
##region FunctionListings

Function Test-ExeOrMsuIsPresent {
<#
.SYNOPSIS
	Check if an .exe or .msu file is present for silent installation/uninstallation. 
.DESCRIPTION
    Check if an .exe or .msu file was found in the same directory as the script or in the Files directory and return true if one was found.
.EXAMPLE
    Test-ExeOrMsuIsPresent
.NOTES
    This function is called to automatically determine if this script should be used to silently install or uninstall an application based on the presence or lack of presence of a .exe or .msu file.
#>
    
    param (
    )

    $result = $true

    if (([string]::IsNullOrEmpty($defaultExeFile)) -and ([string]::IsNullOrEmpty($defaultMsuFile))) {
        $result = $false
    }

    return $result
}

Function Start-Installation {
<#
.SYNOPSIS
	Invoke the appropriate installer function for the detected installer technology.
.DESCRIPTION
    Initiates the installer function appropriate for the installer technology detected using the Find-InstallerTechnology function. 
    The detected installer technology is logged to the default log path.
.PARAMETER deploymentType
	This parameter is passed through this function so that it can be passed to the appropriate installer function.
.EXAMPLE
	Start-Installation -deploymentType Install
.EXAMPLE
	Start-Installation -deploymentType Uninstall
.NOTES
	This function is the main driver for the script and should be called from the Deploy-Application.ps1 script from the PSADT.
#>

    param (
        $deploymentType
    )

    if ($installerTechnology -eq "WindowsInstaller")  {
        Install-UsingWindowsInstaller $defaultExeFile $deploymentType
    }

    if ($installerTechnology -eq "NSIS")  {
        Install-UsingNSISInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Inno Setup") {
        Install-UsingInnoSetupInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "InstallShield") {
        Install-UsingInstallShieldInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "WiXBurn") {
        Install-UsingWixBurnInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Wise")  {
        Install-UsingWiseInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "InstallAWARE")  {
        Install-UsingInstallAWARE $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Install4j")  {
        Install-UsingInstall4j $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "SetupFactory")  {
        Install-UsingSetupFactory $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Office365ClickToRun")  {
        Install-UsingOffice365ClickToRunInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "WindowsUpdateStandaloneInstaller")  {
        Install-UsingWindowsUpdateStandaloneInstaller $defaultMsuFile $deploymentType
    }
}

Function Find-InstallerTechnology {
<#
.SYNOPSIS
	Find which installer technology a .exe or .msu is using.
.DESCRIPTION
    Detects which installer technology a .exe is using by searching the .exe for a string of text unique to that installer (i.e. InstallShield for InstallShield installers).
    For an .msu file the script simply checks for the presence of a .msu file in the script directory or the Files directory.
.EXAMPLE
	Find-InstallerTechnology
.NOTES
    Typically this function is called by the Start-Installation function.
#>
    
    param (
    )

    if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
        Write-Log -Message "Attempting to find which installer technology $appName is using." -Source $appDeployToolkitExtName
    }
        
    ## Load the first 100000 characters of Unicode for the .exe and search it for installer technology reference.
    if (-not ([string]::IsNullOrEmpty($defaultExeFile))) {
        $contentUTF8 = Get-Content -Path $defaultExeFile -Encoding UTF8 -TotalCount 100000
        $contentUnicode = Get-Content -Path $defaultExeFile -Encoding Unicode -TotalCount 100000
    }

    if ((($contentUTF8 -match "Windows installer") -or ($contentUnicode -match "Windows installer")) -and ( -not ($contentUTF8 -match "InstallShield")) -and ( -not ($contentUnicode -match "InstallShield")) -and ( -not ($contentUTF8 -match "ClickToRun")) -and ( -not ($contentUnicode -match "ClickToRun"))) {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Windows installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "WindowsInstaller"
    }

    elseif ((($contentUTF8 -match "NSIS") -or ($contentUnicode -match "NSIS") -or ($contentUTF8 -match "Nullsoft") -or ($contentUnicode -match "Nullsoft")) -and ( -not ($contentUTF8 -match "ClickToRun")) -and ( -not ($contentUnicode -match "ClickToRun")) -and ( -not ($contentUTF8 -match "Inno Setup")) -and ( -not ($contentUnicode -match "Inno Setup")) -and ( -not ($contentUTF8 -match "InstallShield")) -and ( -not ($contentUnicode -match "InstallShield")))  {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the NSIS installer (Nullsoft Scriptable Install System)." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "NSIS"
    }

    elseif ((($contentUTF8 -match "Inno Setup") -or ($contentUnicode -match "Inno Setup")) -and ( -not ($contentUTF8 -match "InstallAWARE")) -and ( -not ($contentUnicode -match "InstallAWARE")))  {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Inno Setup installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "Inno Setup"
    }

    elseif (($contentUTF8 -match "InstallShield") -or ($contentUnicode -match "InstallShield")) {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the InstallShield installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "InstallShield"
    }

    elseif (($contentUTF8 -match "wixburn") -or ($contentUnicode -match "wixburn")) {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Wix Burn installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "WiXBurn"
    }

    elseif (($contentUTF8 -match "WiseMain") -or ($contentUnicode -match "WiseMain") -or ($contentUTF8 -match "Wise Installation") -or ($contentUnicode -match "Wise Installation"))  {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Wise installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "Wise"
    }

    elseif (($contentUTF8 -match "InstallAWARE") -or ($contentUnicode -match "InstallAWARE"))  {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the InstallAWARE installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "InstallAWARE"
    }

    elseif (($contentUTF8 -match "install4j") -or ($contentUnicode -match "install4j"))  {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Install4j installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "Install4j"
    }

    elseif (($contentUTF8 -match "Setup Factory") -or ($contentUnicode -match "Setup Factory"))  {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$userDefinedAppName uses the Setup Factory installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "SetupFactory"
    }

    elseif (($contentUTF8 -match "ClickToRun") -or ($contentUnicode -match "ClickToRun")) {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Microsoft click-to-run installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "Office365ClickToRun"
    }

    elseif (-not ([string]::IsNullOrEmpty($defaultMsuFile))) {
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "$appName uses the Microsoft Windows update standalone installer." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "WindowsUpdateStandaloneInstaller"
    }

    else {
        ## If the installer technology wasn't detected then log that and return to Deploy-Application to execute user defined installation tasks.
        if ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath)) {
            Write-Log -Message "No installer technology was found for $appName." -Source $appDeployToolkitExtName
        }

        $installerTechnology = "Unknown"
    } 

    return $installerTechnology
}

Function Find-UninstallStringInRegistry {
<#
.SYNOPSIS
	Detect the uninstall command for an application in the Windows registry.
.DESCRIPTION
    Search through the HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall key in the registry to find the uninstall command for an application.
    On 32bit systems only this key is searched, on 64bit systems both this key and the WOW6432Node key are searched. The HKCU uninstall key may also be searched.
    The results are logged so if no uninstall string is found the user can modify the uninstall behaviour accordingly.
.EXAMPLE
	Find-UninstallStringInRegistry
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param (
    )

    [string]$uninstallString = ""
    [string]$quietUninstallString = ""
    [hashtable]$returnValues = @{}
    $returnValues.uninstallExe = ""
    $returnValues.arguments = ""

    Write-Log -Message "Attempting to find the uninstall string in the registry." -Source $appDeployToolkitExtName
    
    ## Attempt to find an uninstall string for the app in the Windows uninstall registry key.
    ## Start by checking HKEY_LOCAL_MACHINE, both 32-bit and 64-bit on 64-bit systems (only 32-bit on 32-bit systems).
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall') {
        $uninstallKey = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    if (($Is64Bit) -and (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        $uninstallKey6432 = Get-ChildItem 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    foreach ($key in $uninstallKey) {
        $registryKey = ($key.Name) -replace "HKEY_LOCAL_MACHINE", "HKLM:"
        $registryProperties = Get-ItemProperty -Path $registryKey
        
        if (($registryKey -like "*$appName*") -or ($registryKey -like "*$userDefinedAppName*") -or (($registryProperties.DisplayName) -like "*$appName*") -or (($registryProperties.DisplayName) -like "*$userDefinedAppName*")) {
            if (-not ([string]::IsNullOrEmpty($registryProperties.UninstallString))) {
                $uninstallString = $registryProperties.UninstallString
            }
            
            if (-not ([string]::IsNullOrEmpty($registryProperties.QuietUninstallString))) {
                $quietUninstallString = $registryProperties.QuietUninstallString
            }
        }
    }
    
    ## If no uninstall string was found in the 64-bit uninstall key, look in the 32-bit uninstall key.
    if (-not ([string]::IsNullOrEmpty($uninstallKey6432)) -and ([string]::IsNullOrEmpty($uninstallString)) -and ([string]::IsNullOrEmpty($quietUninstallString))) {
        foreach ($key in $uninstallKey6432) {
            $registryKey = ($key.Name) -replace "HKEY_LOCAL_MACHINE", "HKLM:"
            $registryProperties = Get-ItemProperty -Path $registryKey
         
            if (($registryKey -like "*$appName*") -or ($registryKey -like "*$userDefinedAppName*") -or (($registryProperties.DisplayName) -like "*$appName*") -or (($registryProperties.DisplayName) -like "*$userDefinedAppName*")) {
                if (-not ([string]::IsNullOrEmpty($registryProperties.UninstallString))) {
                    $uninstallString = $registryProperties.UninstallString
                }
            
                if (-not ([string]::IsNullOrEmpty($registryProperties.QuietUninstallString))) {
                    $quietUninstallString = $registryProperties.QuietUninstallString
                }
            }
        }
    }
    
    ## If nothing is found in HKLM, check the current users HKEY_CURRENT_USER uninstall registry key. If the script is running as SYSTEM (as is possible if 
    ## the script has been called by SCCM) then skip the HKCU section.
    if (([Security.Principal.WindowsIdentity]::GetCurrent().Name) -ne "NT AUTHORITY\SYSTEM") {
        if (([string]::IsNullOrEmpty($uninstallString)) -and ([string]::IsNullOrEmpty($quietUninstallString))) {
            if (Test-Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall') {
                $uninstallKey = Get-ChildItem 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
            }

            foreach ($key in $uninstallKey) {
                $registryKey = ($key.Name) -replace "HKEY_CURRENT_USER", "HKCU:"
                $registryProperties = Get-ItemProperty -Path $registryKey
        
                if (-not ([string]::IsNullOrEmpty($registryProperties.UninstallString))) {
                    $uninstallString = $registryProperties.UninstallString
                }
            
                if (-not ([string]::IsNullOrEmpty($registryProperties.QuietUninstallString))) {
                    $quietUninstallString = $registryProperties.QuietUninstallString
                }
            }
        }
    }
    
    ## Return the results in a hashtable so that they can be queried by the requesting uninstall function.
    if (-not ([string]::IsNullOrEmpty($quietUninstallString))) {
        ## If $quietUninstallString does not start with a quote it probably is not encased in them so it should not be split as 
        ## it is unlikely to have arguments.
        if ($quietUninstallString.StartsWith('"')) {
            $uninstallExe, $arguments = $quietUninstallString -split ' +(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)'
        
            if (-not ([string]::IsNullOrEmpty($uninstallExe))) {
                $returnValues.uninstallExe = $uninstallExe.Trim()
            }

            if (-not ([string]::IsNullOrEmpty($arguments))) {
                $returnValues.arguments = $arguments.Trim()
            }
        }

        else {
            $returnValues.uninstallExe = $quietUninstallString.Trim()
        }
    }   
    
    elseif (-not ([string]::IsNullOrEmpty($uninstallString))) {
        ## If $uninstallString does not start with a quote it probably is not encased in them so it should not be split as 
        ## it is unlikely to have arguments.
        if ($uninstallString.StartsWith('"')) {
            $uninstallExe, $arguments = $uninstallString -split ' +(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)'
        
            if (-not ([string]::IsNullOrEmpty($uninstallExe))) {
                $returnValues.uninstallExe = $uninstallExe.Trim()
            }

            if (-not ([string]::IsNullOrEmpty($arguments))) {
                $returnValues.arguments = $arguments.Trim()
            }
        }

        else {
            $returnValues.uninstallExe = $uninstallString.Trim()
        }
    }

    ## If arguments were not encased in quotes the "$returnValues.arguments" may now contain multiple values. This needs to be just one value.
    if ((($returnValues.arguments).Count) -gt 1) {
        $returnValues.arguments = [System.String]::Join(" ", $returnValues.arguments)
    }

    ## If the arguments variable looks like a path, make sure there are quotes around it to protect against paths with spaces in.
    if ($returnValues.arguments -like "*$Env:SystemDrive\*") {
        if ((-not ($returnValues.arguments.StartsWith('"'))) -and (-not ($returnValues.arguments.EndsWith('"')))) {
            $returnValues.arguments = '"' + $returnValues.arguments + '"'
        }
    }

    ## Log the uninstall string and arguments that are found in the registry.
    if ((-not ([string]::IsNullOrEmpty($returnValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($returnValues.arguments)))) {
        Write-Log -Message "$appName uninstallation string found in registry. It is: $($returnValues.uninstallExe) $($returnValues.arguments)" -Source $appDeployToolkitExtName        
    }
        
    elseif (-not ([string]::IsNullOrEmpty($returnValues.uninstallExe))) {
        Write-Log -Message "$appName uninstallation string found in registry. It is: $($returnValues.uninstallExe)" -Source $appDeployToolkitExtName            
    }
        
    else {
        Write-Log -Message "$appName uninstallation string could not be found in the registry." -Source $appDeployToolkitExtName
    }
    
    return $returnValues
}

Function Install-UsingWindowsInstaller {
<#
.SYNOPSIS
	Silently installs an .exe that uses the Windows installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install a Windows installer application and starts the installation. 
    For uninstalls this application uses the uninstall command found in the registry.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingWixBurnInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingWixBurnInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = '/qn'
    }

    if ($deploymentType -eq "Uninstall") {
        $uninstallValues = Find-UninstallStringInRegistry

        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
        }
        
        elseif (-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) {
            $defaultExeFile, $arguments = $uninstallValues.uninstallExe -split "\s\s|\s"
        }
        
        else {
            Write-Log -Message "Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }

        ## Add the log switch /l*v to verbosely log the uninstall to the user specified log location.
        $arguments = $arguments + ' /l*v ' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'

        ## If UninstallString specified MsiExec.exe /I, replace with /X as this is the proper uninstall switch.
        if ($arguments -like "*/I*") {
            $arguments = $arguments -replace ("/I", "/X")
        }

        ## Add the silent switch /qn if it is not included in the arguments obtained from the registry.
        if (-not ($arguments -like "*/qn*")) {
            $arguments = $arguments + " /qn"
        }
    }

    ## Install or uninstall string has been worked out, execution process.
    ## Log file is sometimes written to %TEMP% during an install or uninstall.
    ## If this application has logged to %TEMP%, the Log file will be copied to the user defined log location when the install finishes.
    [DateTime]$installationStartTime = Get-Date
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait

    if ($deploymentType -eq "Install") {
        $logFile = Get-ChildItem -Path $env:TEMP -Filter "*.txt" | Where-Object {$_.LastWriteTime -gt $installationStartTime}
        if ($logFile -is [Array]){
            $logCounter = 1
            foreach ($log in $logFile) {
                Copy-Item -Path $log.FullName -Destination ("$configToolkitLogDir\$appExeLogFileName" + "_" + "$logCounter" + '_Install.log')
                $logCounter++
            }
        }

        else {
            if (-not ([string]::IsNullOrEmpty($logFile.FullName))) {
                Copy-Item -Path $logFile.FullName -Destination ("$configToolkitLogDir\$appExeLogFileName" + '_Install.log')
            }
        }
    }
}

Function Install-UsingNSISInstaller {
<#
.SYNOPSIS
	Silently installs an .exe that uses the NSIS installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install an NSIS application and starts the installation.
    For uninstalls this application uses the uninstall command found in the registry. NSIS does not output a log file so no log is written. 
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingNSISInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingNSISInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>
    
    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = "/AllUsers /S -silent --allusers=1"
    }

    if ($deploymentType -eq "Uninstall") {
        $uninstallValues = Find-UninstallStringInRegistry
        
        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments

            ## Add the uninstall switch -uninstall if it is not included in the arguments obtained from the registry.
            if (-not ($arguments -like "*/uninstall*") -or (-not ($arguments -like "*-uninstall*"))) {
                $arguments = $arguments + " -uninstall"
            }

            ## Add the all users switch /AllUsers if it is not included in the arguments obtained from the registry.
            if (-not ($arguments -like "*/AllUsers*")) {
                $arguments = $arguments + " /AllUsers --allusers=1"
            }

            ## Add the silent switch /S if it is not included in the arguments obtained from the registry.
            if (-not ($arguments -like "*/S*")) {
                $arguments = $arguments + " /S -silent"
            }
        }
        
        elseif (-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = "-uninstall /AllUsers --allusers=1 /S -silent"    
        }
        
        else {
            Write-Log -Message "Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }
    
    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingInnoSetupInstaller {
<#
.SYNOPSIS
	Silently installs an .exe that uses the Inno Setup installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install an NSIS application and starts the installation. Inno Setup supports .inf files with pre-configured settings (such as install path), so if a .inf file is found this is used in the install.
    For uninstalls this application uses the uninstall command found in the registry. Both install and uninstall log files are written.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingInnoSetupInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingInnoSetupInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>
    
    param(
        $defaultExeFile,
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        [string]$infFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetExtension($_.Name) -eq '.inf') } | Select-Object -ExpandProperty 'FullName' -First 1
        if ([string]::IsNullOrEmpty($infFile)) {
            $arguments = '/sp- /verysilent /norestart /SuppressMSGBoxes /LOG=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
        }

        else {
            Write-Log -Message "Configuration .inf file found: $infFile." -Source $appDeployToolkitExtName
            $arguments = '/sp- /verysilent /norestart /SuppressMSGBoxes /LoadInf=' + '"' + $infFile + '"' + ' /LOG=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
        }
    }

    if ($deploymentType -eq "Uninstall") {
        ## Inno Setup sometimes places an uninstaller in the install directory of the app, attempt to find this in the registry and then uninstall.
        $uninstallValues = Find-UninstallStringInRegistry
        
        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments

            ## If the /silent switch is specified in the registry, replace with /verysilent as this hides the uninstall status screen as well.
            if ($arguments -like "*/silent*") {
                $arguments = $arguments -replace ("/silent", "/verysilent")
            }

            ## Add the silent switch /verysilent if it is not included in the arguments obtained from the registry.
            if (-not ($arguments -like "*/verysilent*")) {
                $arguments = $arguments + " /verysilent"
            }

            ## Add the no restart switch /norestart if it is not included in the arguments obtained from the registry.
            if (-not ($arguments -like "*/norestart*")) {
                $arguments = $arguments + " /norestart"
            }

            ## Add the log switch /LOG if it is not included in the arguments obtained from the registry.
            if (-not ($arguments -like "*/LOG*")) {
                $arguments = $arguments + ' /LOG=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'
            }
        }
        
        elseif (-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = '/verysilent /norestart /SuppressMSGBoxes /LOG=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'   
        }
        
        else {
            $arguments = '/uninstall /verysilent /norestart /SuppressMSGBoxes /LOG=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'
        }
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingInstallShieldInstaller {
<#
.SYNOPSIS
	Silently installs an .exe that uses the InstallShield installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install an InstallShield application and starts the installation. 
    InstallShield requires .iss files for both silent install and uninstall so this is checked and added to the install/uninstall string. 
    The uninstaller uses the same .exe as the installer but with different arguments.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingInstallShieldInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingInstallShieldInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        [string]$installISS = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'install.iss') } | Select-Object -ExpandProperty 'FullName'
        if (-not ([string]::IsNullOrEmpty($installISS))) {
            Write-Log -Message "Install .iss file found: $installISS." -Source $appDeployToolkitExtName
            $arguments = '/s /SMS /f1' + '"' + $installISS + '"' + ' /f2' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
        }
        
        else {
            ## If a installation .iss file is not available log the error and end the installation.
            Write-Log -Message "$appName installation .iss file not found. Install will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    if ($deploymentType -eq "Uninstall") {
        [string]$uninstallISS = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'uninstall.iss') } | Select-Object -ExpandProperty 'FullName'
        if (-not ([string]::IsNullOrEmpty($uninstallISS))) {
            Write-Log -Message "Uninstall .iss file found: $uninstallISS." -Source $appDeployToolkitExtName
            $arguments = '/s /SMS /uninst /f1' + '"' + $uninstallISS + '"' + ' /f2' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'
        }
        
        else {
            ## If a uninstallation .iss file is not available log the error and end the uninstallation.
            Write-Log -Message "$appName uninstallation .iss file not found. Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingWixBurnInstaller {
<#
.SYNOPSIS
	Silently installs an .exe that uses the WiX Burn installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install a WiX Burn application and starts the installation. 
    The uninstaller uses the same .exe as the installer but with different arguments.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingWixBurnInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingWixBurnInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = '/s /norestart /l ' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
    }

    if ($deploymentType -eq "Uninstall") {
        $arguments = '/s /norestart /uninstall /l ' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingWiseInstaller {
<#
.SYNOPSIS
	Silently installs an .exe that uses the Wise installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install a Wise application and starts the installation. 
    For uninstalls this application uses the uninstall command found in the registry. Wise does not output a log file, so no log is written.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingWiseInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingWiseInstaller -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = '/S'
    }

    if ($deploymentType -eq "Uninstall") {
        $uninstallValues = Find-UninstallStringInRegistry

        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
        }
        
        elseif (-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) {
            $uninstallStringExe, $blankSpace, $uninstallStringLog = [regex]::matches(($uninstallValues.uninstallExe),'(?<=\").+?(?=\")').Value
            $defaultExeFile = $uninstallStringExe
            $arguments = '/S /Z ' + '"' + $uninstallStringLog + '"'
        }
        
        else {
            Write-Log -Message "Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingInstallAWARE {
<#
.SYNOPSIS
	Silently installs an .exe that uses the InstallAWARE installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install an InstallAWARE application and starts the installation. 
    The uninstaller uses the same .exe as the installer but with different arguments.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingInstallAWARE -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingInstallAWARE -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = '/s /l=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
    }

    if ($deploymentType -eq "Uninstall") {
        $arguments = '/s MODIFY=FALSE REMOVE=TRUE UNINSTALL=YES /l=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingInstall4j {
<#
.SYNOPSIS
	Silently installs an .exe that uses the Install4j installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install an Install4j application and starts the installation. 
    For uninstalls this application uses the uninstall executable found in the registry, which is combined with silent switches and log file location.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingInstall4j -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingInstall4j -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = '-q -Dinstall4j.keepLog=true -Dinstall4j.alternativeLogfile=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
    }

    if ($deploymentType -eq "Uninstall") {
        $uninstallValues = Find-UninstallStringInRegistry

        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
        }
        
        elseif (-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) {
            $uninstallStringExe, $blankSpace, $uninstallStringLog = [regex]::matches(($uninstallValues.uninstallExe),'(?<=\").+?(?=\")').Value
            $defaultExeFile = $uninstallStringExe
            $arguments = '-q -Dinstall4j.keepLog=true -Dinstall4j.alternativeLogfile=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log' + '"'
        }
        
        else {
            Write-Log -Message "Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingSetupFactory {
<#
.SYNOPSIS
	Silently installs an .exe that uses the Setup Factory installer technology.
.DESCRIPTION
    Sets the arguments needed to silently install an Install4j application and starts the installation. 
    For uninstalls this application uses the uninstall executable found in the registry, which is combined with silent switches and log file location.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingInstall4j -defaultExeFile "C:\Product\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingInstall4j -defaultExeFile "C:\Product\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        $arguments = '/S /NOINIT'
    }

    if ($deploymentType -eq "Uninstall") {
        ## Setup Factory installers always have "Setup Factory Runtime" as the product name rather than the actual product details. Therefore user defined app name needs to be used.
        $appName = $userDefinedAppName
        $uninstallValues = Find-UninstallStringInRegistry

        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments + " /S"
        }
        
        else {
            Write-Log -Message "Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    ## Install or uninstall string has been worked out, execution process
    ## Log file is written to %TEMP% (no way to specify alternative location with Setup Factory)
    ## Log file will be copied to user defined log location when the install / uninstall finishes.

    [DateTime]$installationStartTime = Get-Date
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait

    $logFile = Get-ChildItem -Path $env:TEMP -Filter "*Log.txt" | Where-Object {$_.LastWriteTime -gt $installationStartTime}
    if ($deploymentType -eq "Install") {
        Copy-Item -Path $logFile.FullName -Destination ("$configToolkitLogDir\$appExeLogFileName" + '_Install.log')
    }

    if ($deploymentType -eq "Uninstall") {
        Copy-Item -Path $logFile.FullName -Destination ("$configToolkitLogDir\$appExeLogFileName" + '_Uninstall.log')
    }

}

Function Install-UsingOffice365ClickToRunInstaller {
<#
.SYNOPSIS
	Silently installs Microsoft Office 365 click-to-run.
.DESCRIPTION
    This is one of the special functions designed to install a specific product.
    It sets the arguments needed to silently install Office 365, which requires the presence of an .xml file which configures the product.
    For more information see this document from Microsoft: https://docs.microsoft.com/en-us/deployoffice/deploy-office-365-proplus-from-a-local-source
    For uninstalls this application uses the uninstall command found in the registry.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingOffice365ClickToRunInstaller -defaultExeFile "C:\Office365\Setup.exe" -deploymentType Install
.EXAMPLE
	Install-UsingOffice365ClickToRunInstaller -defaultExeFile "C:\Office365\Setup.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    if ($deploymentType -eq "Install") {
        [string]$xmlFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetExtension($_.Name) -eq '.xml') } | Select-Object -ExpandProperty 'FullName' -First 1
        if (-not ([string]::IsNullOrEmpty($xmlFile))) {
            Write-Log -Message "Configuration .xml file found: $xmlFile." -Source $appDeployToolkitExtName
            $arguments = '/configure ' + '"' + $xmlFile + '"'
        }

        else {
            ## If a configuration .xml file is not available log the error and end the installation.
            Write-Log -Message "Microsoft Office 365 configuration .xml file not found. Install will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    ## For Office 365 uninstall look for O365ProPlusRetail key in registry for uninstall string.
    if ($deploymentType -eq "Uninstall") {
        $realAppName = $appName
        $appName = "O365ProPlusRetail"
        $uninstallValues = Find-UninstallStringInRegistry
        
        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
        }

        else {
            Write-Log -Message "Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
        
        $appName = $realAppName
        $arguments = $arguments + ' DisplayLevel=False'
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Install-UsingWindowsUpdateStandaloneInstaller {
<#
.SYNOPSIS
	Silently installs Microsoft Windows update files.
.DESCRIPTION
    This is one of the special functions designed to install a specific product.
    It sets the arguments needed to silently install a Windows update.
    The uninstaller uses the same .msu as the installer but with different arguments. Windows update install logs are written in .evtx format which can be opened in Windows Event Viewer.
.PARAMETER defaultExeFile
	The full path to the executable installer or uninstaller.
.PARAMETER deploymentType
	Passed from the Start-Installation function and used here to determine if the user wants to install or uninstall the application.
.EXAMPLE
	Install-UsingWindowsUpdateStandaloneInstaller -defaultExeFile "wusa.exe" -deploymentType Install
.EXAMPLE
	Install-UsingWindowsUpdateStandaloneInstaller -defaultExeFile "wusa.exe" -deploymentType Uninstall
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param(
        $defaultExeFile, 
        $deploymentType
    )

    $defaultExeFile = "wusa.exe"

    if ($deploymentType -eq "Install") {
        $arguments = '$defaultMsuFile /quiet /norestart /log:' + '"' + "$configToolkitLogDir\$appMsuLogFileName"  + '_Install.evtx' + '"'
    }

    if ($deploymentType -eq "Uninstall") {
        $arguments = '$defaultMsuFile /uninstall /quiet /norestart /log:' + '"' + "$configToolkitLogDir\$appMsuLogFileName"  + '_Uninstall.evtx' + '"'
    }

    ## Install or uninstall string has been worked out, execution process
    Write-Log -Message "Full command is $defaultExeFile $arguments" -Source $appDeployToolkitExtName
    Start-Process $defaultExeFile -ArgumentList $arguments -Wait
}

Function Test-SupportedInstallerType {
<#
.SYNOPSIS
    Allows a user of this script to test whether or not their application will be supported.
.DESCRIPTION 
    This function is used when a user runs this script and specifies the TestSupportedInstallerTypePath variable at the command line,
    along with a path to an installer executable.
    It will output the installer type that is used if it is known by the script, or inform the user that this script does not natively
    support the installer.
.EXAMPLE
	ZeroConfigExeInstallation.ps1 -TestSupportedInstallerTypePath "C:\Temp\Application\setup.exe"
.NOTES
	This is an internal script function and should typically not be called directly. It should be used as shown in the example.
#>
    param (
    )

    ## Figure out what installer technology is used by the specified executable.
    $defaultExeFile = $TestSupportedInstallerTypePath
    [string]$installerTechnology = Find-InstallerTechnology

    ## Write the result to the console.
    if ($installerTechnology -eq "WindowsInstaller")  {
        Write-Host "This executable uses the Windows Installer and is supported by this script."
    }

    if ($installerTechnology -eq "NSIS")  {
        Write-Host "This executable uses the NSIS installer and is supported by this script."
    }

    If ($installerTechnology -eq "Inno Setup") {
        Write-Host "This executable uses the Inno Setup installer and is supported by this script."
    }

    If ($installerTechnology -eq "InstallShield") {
        Write-Host "This executable uses the InstallShield installer and is supported by this script."
    }

    If ($installerTechnology -eq "WiXBurn") {
        Write-Host "This executable uses the WiXBurn installer and is supported by this script."
    }

    If ($installerTechnology -eq "Wise")  {
        Write-Host "This executable uses the Wise installer and is supported by this script."
    }

    If ($installerTechnology -eq "InstallAWARE")  {
        Write-Host "This executable uses the InstallAWARE installer and is supported by this script."
    }

    If ($installerTechnology -eq "Install4j")  {
        Write-Host "This executable uses the Install4j installer and is supported by this script."
    }

    If ($installerTechnology -eq "SetupFactory")  {
        Write-Host "This executable uses the Setup Factory installer and is supported by this script."
    }

    If ($installerTechnology -eq "Office365ClickToRun")  {
        Write-Host "This executable uses the Office 365 click-to-run installer and is supported by this script."
    }

    If ($installerTechnology -eq "WindowsUpdateStandaloneInstaller")  {
        Write-Host "This executable uses the Windows Update Standalone installer and is supported by this script."
    }

    If ($installerTechnology -eq "Unknown")  {
        Write-Host "The installer used by this executable is not supported by this script."
    }
}

##endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
##region ScriptBody

## If this script has been started with the $TestSupportedInstallerTypePath variable set, run the Test-SupportedInstallerType function and end the running of the script.
## This must run here to prevent other variables or functions running unnecessarily.
if (-not ([string]::IsNullOrEmpty($TestSupportedInstallerTypePath))) {
    Test-SupportedInstallerType

    Exit
}

## Find which installer technology is being used by the installer.
[string]$installerTechnology = Find-InstallerTechnology
if (-not ([string]::IsNullOrEmpty($defaultExeFile))) {
    
    ## InstallShield installer files sometimes contain info about InstallShield rather than the product. Get this information from the install.iss file instead.
    if ($installerTechnology -eq "InstallShield") {
        [string]$installISS = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'install.iss') } | Select-Object -ExpandProperty 'FullName'
        
        if (-not ([string]::IsNullOrEmpty($installISS))) {
            foreach ($line in [System.IO.File]::ReadLines($installISS)) {
                if ($line -like "*Company=*") { 
                    $appVendor = $line -replace "Company=",""
                }

                if ($line -like "*Name=*") { 
                    $appName = $line -replace "Name=",""
                }

                if ($line -like "*Version=*") { 
                    $appVersion = $line -replace "Version=",""
                }
            }
        }
    }

    ## Wise installer files sometimes do not contain product vender, name or version information, but do have some info in the description, so using that instead if it's available.
    elseif ($installerTechnology -eq "Wise") {
        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.FileDescription))) {
            $fileDescription = (Get-Item -Path $defaultExeFile).VersionInfo.FileDescription
            $appName = $fileDescription
        }
    }

    ## Office 365 click-to-run version is not held in the setup.exe file, but in the folder name Office\Data\[version].
    ## Purposefully leave out AppVendor as Microsoft uses their name in the app name (so it doesn't read Microsoft Coproration Microsoft Office 2016...).
    elseif ($installerTechnology -eq "Office365ClickToRun") {
        $appName =  (Get-Item -Path $defaultExeFile).VersionInfo.FileDescription.Trim()
        $appVersion = (Get-ChildItem -Path "$dirFiles\Office\Data\" -Recurse | ?{ $_.PSIsContainer }).Name
    }

    ## If none of the conditions above are met, use the details in the .exe to fill in $appVendor, $appName and $appVersion.
    else {

        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.CompanyName))) {
            $appVendor = (Get-Item -Path $defaultExeFile).VersionInfo.CompanyName.Trim()
        }
        
        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.ProductName))) {       
            $appName =  (Get-Item -Path $defaultExeFile).VersionInfo.ProductName.Trim()
        }

        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.ProductVersion))) {
            $appVersion = (Get-Item -Path $defaultExeFile).VersionInfo.ProductVersion.Trim()
        }
    }

    ## Remove any special characters in the app name for use in the log name. Also remove the words "installer", "install", "installation or "setup" if they are found in the app name.
    $appName = ($appName -replace "installer", "").Trim()
    $appName = ($appName -replace "install", "").Trim()
    $appName = ($appName -replace "installation", "").Trim()
    $appName = ($appName -replace "setup", "").Trim()
    $appExeLogFileName = ($appName -replace '[\W]', '')

    Write-Log -Message "App Vendor [$appVendor]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Name [$appName]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Version [$appVersion]." -Source $appDeployToolkitExtName
}

## Msu installers don't contain update information in the file properties, so just use Microsoft as manufacturer and the KB number as the name.
if (-not ([string]::IsNullOrEmpty($defaultMsuFile))) {
    $appVendor = "Microsoft"
    $defaultMsuFile -match '\b[a-zA-Z]{2}\d*\b' | Out-Null
    $appName =  $Matches[0]
    $appMsuLogFileName = ($appName -replace '[\W]', '')

    Write-Log -Message "App Vendor [$appVendor]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Name [$appName]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Version [$appVersion]." -Source $appDeployToolkitExtName
}

# SIG # Begin signature block
# MIIdZAYJKoZIhvcNAQcCoIIdVTCCHVECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVOaEqtO4gtmvX1z2f2u0EuC/
# f+6gghiIMIIFTDCCBDSgAwIBAgIRAKLa/6xNrUXkkS75zMNjpi0wDQYJKoZIhvcN
# AQELBQAwfDELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3Rl
# cjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSQw
# IgYDVQQDExtTZWN0aWdvIFJTQSBDb2RlIFNpZ25pbmcgQ0EwHhcNMTkxMDMxMDAw
# MDAwWhcNMjIxMDMwMjM1OTU5WjCBkjELMAkGA1UEBhMCR0IxEDAOBgNVBBEMB0VD
# MkE0TkUxDzANBgNVBAgMBkxvbmRvbjEaMBgGA1UECQwRODYtOTAgUGF1bCBTdHJl
# ZXQxITAfBgNVBAoMGEsgU3RyZWV0IENvbnN1bHRhbmN5IEx0ZDEhMB8GA1UEAwwY
# SyBTdHJlZXQgQ29uc3VsdGFuY3kgTHRkMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
# MIIBCgKCAQEAu9lGsqAElNe5xGFWxK19zB0dl3R81R/4VdMh212QjeNvGb7kZvnn
# BuAHeQgdip3BIdwp6dpcxlhDKCfWk2BOgpl9iyCsY1H/cu95nL+tOHUsLieow8m+
# JXfu17Tqta+EvTyi1oau/6VVRLpF/DQS1R1dhO1PkQtrxJ3wWPhOaf5IgeRUA+vR
# UI2cIkYs4XFoUQ23xnM3nDWCVJmviFk+xTjckLf9azmyjb1C6rCvB5FEtlrwHaoM
# ImJGtPJfwwRwopyfkr8KrKYyqsz/+Vua/dRihQNys9/X7zHnGhU8+2RXfkQrJ97P
# X5Bxvv7KY6JHSja21G8mJK6E8XO/QPAK8QIDAQABo4IBsDCCAawwHwYDVR0jBBgw
# FoAUDuE6qFM6MdWKvsG7rWcaA4WtNA4wHQYDVR0OBBYEFH3GRaH18DbrtLe22E9y
# S2I5UOkaMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMBEGCWCGSAGG+EIBAQQEAwIEEDBABgNVHSAEOTA3MDUGDCsGAQQB
# sjEBAgEDAjAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQUzBD
# BgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29S
# U0FDb2RlU2lnbmluZ0NBLmNybDBzBggrBgEFBQcBAQRnMGUwPgYIKwYBBQUHMAKG
# Mmh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1JTQUNvZGVTaWduaW5nQ0Eu
# Y3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAoBgNVHREE
# ITAfgR1jb2Rlc2lnbmluZ0BrZXZpbnN0cmVldC5jby51azANBgkqhkiG9w0BAQsF
# AAOCAQEAKB8JPeWA4wG/BXbLTWY1oPAnpgrKPBObNcQUFxmMCiGo3WLYW+1GCqKi
# CLvG0DB74dwOphxvLPAKj1gpn3YEUrO02Kvb1Y0LJthRQOUzr6LZpZTHONgwINlt
# q3Fuu+LCnK+2mf7W0cBmS2AKD0pyZaJvLded6mwftYcbTp+xxu1b+gM/fsomjd/Y
# YLQUzAqXbY2q2g7toeuNVnyMnfMYmMvpN2VFWG8zZd8SBDH2/g2NcM5POz5HjNgq
# +bMhkGfoZjWhm2NPkpwJTnqKnpfLQrtLstMwnWQhYWB3fY+q0V9CVYtfBxzdqKb5
# SGNGMdVh/sldL1W0ub1CNNRkk6sKzTCCBfUwggPdoAMCAQICEB2iSDBvmyYY0ILg
# ln0z02owDQYJKoZIhvcNAQEMBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpO
# ZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UEChMVVGhlIFVT
# RVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNBIENlcnRpZmlj
# YXRpb24gQXV0aG9yaXR5MB4XDTE4MTEwMjAwMDAwMFoXDTMwMTIzMTIzNTk1OVow
# fDELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4G
# A1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSQwIgYDVQQD
# ExtTZWN0aWdvIFJTQSBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCGIo0yhXoYn0nwli9jCB4t3HyfFM/jJrYlZilAhlRGdDFi
# xRDtsocnppnLlTDAVvWkdcapDlBipVGREGrgS2Ku/fD4GKyn/+4uMyD6DBmJqGx7
# rQDDYaHcaWVtH24nlteXUYam9CflfGqLlR5bYNV+1xaSnAAvaPeX7Wpyvjg7Y96P
# v25MQV0SIAhZ6DnNj9LWzwa0VwW2TqE+V2sfmLzEYtYbC43HZhtKn52BxHJAteJf
# 7wtF/6POF6YtVbC3sLxUap28jVZTxvC6eVBJLPcDuf4vZTXyIuosB69G2flGHNyM
# fHEo8/6nxhTdVZFuihEN3wYklX0Pp6F8OtqGNWHTAgMBAAGjggFkMIIBYDAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQUDuE6qFM6MdWK
# vsG7rWcaA4WtNA4wDgYDVR0PAQH/BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAw
# HQYDVR0lBBYwFAYIKwYBBQUHAwMGCCsGAQUFBwMIMBEGA1UdIAQKMAgwBgYEVR0g
# ADBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLnVzZXJ0cnVzdC5jb20vVVNF
# UlRydXN0UlNBQ2VydGlmaWNhdGlvbkF1dGhvcml0eS5jcmwwdgYIKwYBBQUHAQEE
# ajBoMD8GCCsGAQUFBzAChjNodHRwOi8vY3J0LnVzZXJ0cnVzdC5jb20vVVNFUlRy
# dXN0UlNBQWRkVHJ1c3RDQS5jcnQwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVz
# ZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIBAE1jUO1HNEphpNveaiqMm/EA
# AB4dYns61zLC9rPgY7P7YQCImhttEAcET7646ol4IusPRuzzRl5ARokS9At3Wpwq
# QTr81vTr5/cVlTPDoYMot94v5JT3hTODLUpASL+awk9KsY8k9LOBN9O3ZLCmI2pZ
# aFJCX/8E6+F0ZXkI9amT3mtxQJmWunjxucjiwwgWsatjWsgVgG10Xkp1fqW4w2y1
# z99KeYdcx0BNYzX2MNPPtQoOCwR/oEuuu6Ol0IQAkz5TXTSlADVpbL6fICUQDRn7
# UJBhvjmPeo5N9p8OHv4HURJmgyYZSJXOSsnBf/M6BZv5b9+If8AjntIeQ3pFMcGc
# TanwWbJZGehqjSkEAnd8S0vNcL46slVaeD68u28DECV3FTSK+TbMQ5Lkuk/xYpMo
# JVcp+1EZx6ElQGqEV8aynbG8HArafGd+fS7pKEwYfsR7MUFxmksp7As9V1DSyt39
# ngVR5UR43QHesXWYDVQk/fBO4+L4g71yuss9Ou7wXheSaG3IYfmm8SoKC6W59J7u
# mDIFhZ7r+YMp08Ysfb06dy6LN0KgaoLtO0qqlBCk4Q34F8W2WnkzGJLjtXX4oemO
# CiUe5B7xn1qHI/+fpFGe+zmAEc3btcSnqIBv5VPU4OOiwtJbGvoyJi1qV3AcPKRY
# LqPzW0sH3DJZ84enGm1YMIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr1tXq5hfwZjAN
# BgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2Vy
# dCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQxMDIyMDAwMDAw
# WjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERp
# Z2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnmfGt/a4ydVfiS
# 457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk6qhLLJGJzF4o
# 9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxVh3lIRvfKDo2n
# 3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBMMWSN4+v6GYeo
# fs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD6dpAoVk62RUJ
# V5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1MIIDMTAOBgNVHQ8B
# Af8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCC
# Ab8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6C
# AVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBp
# AGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABh
# AG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBD
# AFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5
# ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABs
# AGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABv
# AHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBj
# AGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStn
# As0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0GA1UdHwR2MHQwOKA2
# oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENB
# LTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZI
# hvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x1GosMe06FxlxF82p
# G7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm6mlXIV00Lx9xsIOU
# GQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dgc6aSwNOOMdgv
# 420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu19aDxxncGKBXp
# 2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHpigtt7BIYvfdVVEAD
# kitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbNMIIFtaADAgECAhAG/fkD
# lgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAi
# BgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAw
# MDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERp
# Z2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
# AQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xNRZFcgZ/tLJz4Flnf
# nrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2CAOrDDT+GEmC/sfH
# MUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIhFonGcIj5BZd9o8dD
# 3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/JksUkK5ZZgrEjb7S
# zgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt1/HYzx4KdFxC
# uGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2MA4GA1UdDwEB/wQE
# AwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIGCCsGAQUFBwMDBggr
# BgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCCAbQGCmCGSAGG/WwA
# AQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wt
# Y3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAg
# AHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAg
# AGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABv
# AGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBu
# AGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBl
# AGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBs
# AGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBk
# ACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCG
# SAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggr
# BgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdo
# dHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290
# Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3Js
# NC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1Ud
# DgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF66Kv9JLLgjEt
# UYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4HbZbKa9Sz1Lp
# MUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKtOHisdV0XFzRyR4WU
# VtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0aqRRbpoYxYqio
# M+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5vEwSV/5f4R68Al2o
# /vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzunEIOz5HXJ7cW7g/D
# vXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1PL1qVCCkQJjGC
# BEYwggRCAgEBMIGRMHwxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1h
# bmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEkMCIGA1UEAxMbU2VjdGlnbyBSU0EgQ29kZSBTaWduaW5nIENBAhEAotr/
# rE2tReSRLvnMw2OmLTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUniI7C4qWMcfflnG9vKL/cXFw
# wlowDQYJKoZIhvcNAQEBBQAEggEAqAIgXaGnfCLk7j/2gQhosVXV1sgddwAtKX7+
# QEB+k+SgCB9x4vA1cs3NVG7sHU2ORT/a2AljKOfWh2qmO+tiAiKfh1HvjI2evALi
# A1A6vpJC0uyoKcztFkzHBEaojuw1O/pNQupj4CyBGSHnril6tn7Dfi403QUm6wj2
# LAWV3dwRwaw5+9NIrVy5RjhGOz8Wc//ZhDsXu5AGiAxhyeM8iOX0CJeO5y0k3L8j
# yaLr4rfqdV70kJKifEP/h3Jr4ABv0KwyDUfRebLfEUIhKDhBbLdSZ8GZHIqPP+NN
# PLvJYDoaNK+ocOfgV8VOUFLeB2gBr8WsKQuNlP7mostuZ9Rt1qGCAg8wggILBgkq
# hkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMT
# GERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfwZjAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjAwNjMwMTkyMTM0WjAjBgkqhkiG9w0BCQQxFgQUmg0lEnzE+F9oaWrH5sHz
# OUMfj28wDQYJKoZIhvcNAQEBBQAEggEAHWUgr3CaZ5hBgGDcbhA2eHioqFOBtQKw
# 7B8tiDGq3H+/S9NkHXuGcxmrCc1bNL4gy/7FZEn21Nu+9hS8b5Cp61JhBZk+BLci
# Y7F1uK5ubQsIcU69075kqsPEV4cIp2A20+7fhOufzBcMDczMR2jhELBWvG7PZyY4
# dX5wU2hknyeTYa3tEi0S073bvlUpp0flEC68sPbUzO/6Yg5yLBajsCUZSTj3OSf4
# qAsezEm1QKcxgu/w3io+tszGNE4Ey/UbeIkNwf9ZCkDfydaEjFyAG7JFMIyVz6A5
# xXNoHCIb+ir5cKIumXg4MjSp08fNnpRSPdskOMfSmnk6LiejMekmWw==
# SIG # End signature block
