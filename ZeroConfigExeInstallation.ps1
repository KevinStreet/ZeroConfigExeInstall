<#
.SYNOPSIS
	This script is used to determine which installer an executable is using and silently install or uninstall it using default parameters.
	# LICENSE #
	Zero-Config Executable Installation. 
	Copyright (C) 2019 - Kevin Street.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. 
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
    This script attempts to find which installer technology an software product is using and then installdes it silently using the default parameters.
    The script supports the following installer products:
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
    This paramter is passed to the script when invoked from the PADT.
.EXAMPLE
    ZeroConfigExeInstallation.ps1 -deploymentType Install
.EXAMPLE
    ZeroConfigExeInstallation.ps1 -deploymentType Uninstall
.NOTES
    Script version: 0.3.0
    Release date: 02/06/2019.
    Author: Kevin Street.
.LINK
	https://kevinstreet.co.uk

#Requires -Version 2.0
#>

[CmdletBinding()]

param (

)

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
##region VariableDeclaration

[string]$appDeployToolkitExtName = 'ZeroConfigExe'
[string]$appDeployExtScriptFriendlyName = 'Zero-Config Executable Installation'
[version]$appDeployExtScriptVersion = [version]'0.3.0'
[string]$appDeployExtScriptDate = '02/06/2019'

## Check for Exe installer and modify the installer path accordingly.
## If multiple .exe files are found attempt to find setup.exe or install.exe and use those. If neither exist the user must specify the installer .exe in $installerExecutable in Deploy-Application.ps1.
if ([string]::IsNullOrEmpty($installerExecutable)) {
    [array]$exesInPath = (Get-ChildItem -Path "$dirFiles\*.exe").Name
    if ($exesInPath.Count -gt 1) {
        if ($exesInPath -contains "setup.exe") {
            [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'setup.exe') } | Select-Object -ExpandProperty 'FullName'
        }

        elseif ($exesInPath -contains "install.exe") {
            [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq 'install.exe') } | Select-Object -ExpandProperty 'FullName'
        }

        else {
            Write-Log -Message "Multiple .exe files found but not sure which to use as the install. The installer .exe must be specified in Deploy-Application.ps1." -Source $appDeployToolkitExtName
        }
    }

    else {
        [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetExtension($_.Name) -eq '.exe') } | Select-Object -ExpandProperty 'FullName' -First 1
    }
}

else {
    [string]$defaultExeFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq "$installerExecutable") } | Select-Object -ExpandProperty 'FullName'
}

## Check for Msu installer and modify the installer path accordingly.
## If multiple .msu files are found iform the user they must specify which one they want in $installerExecutable in Deploy-Application.ps1.
if ([string]::IsNullOrEmpty($defaultExeFile)) {
    if ([string]::IsNullOrEmpty($installerExecutable)) {
        [array]$msusInPath = (Get-ChildItem -Path "$dirFiles\*.msu").Name
        if ($msusInPath -gt 1) {
            Write-Log -Message "Multiple .msu files found but not sure which one to use. Please reduce to one .msu or specify which .msu to use in Deploy-Application.ps1." -Source $appDeployToolkitExtName
        }

        else {
            [string]$defaultMsuFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetExtension($_.Name) -eq '.msu') } | Select-Object -ExpandProperty 'FullName' -First 1
        }
    }

    else {
        [string]$defaultMsuFile = Get-ChildItem -LiteralPath $dirFiles -ErrorAction 'SilentlyContinue' | Where-Object { (-not $_.PsIsContainer) -and ([IO.Path]::GetFileName($_.Name) -eq "$installerExecutable") } | Select-Object -ExpandProperty 'FullName'
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
    This function is called to automatically determine if this script should be use to silently install or uninstall an application based on the presence or lack of presence of a .exe or .msu file.
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
	This function is the main driver for the script and should be called from the Deploy-Application.ps1 script from the PADT.
#>

    param (
        $deploymentType
    )

    if ($installerTechnology -eq "NSIS")  {
        Write-Log -Message "$appName uses the NSIS installer (Nullsoft Scriptable Install System)." -Source $appDeployToolkitExtName
        Install-UsingNSISInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Inno Setup") {
        Write-Log -Message "$appName uses the Inno Setup installer." -Source $appDeployToolkitExtName
        Install-UsingInnoSetupInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "InstallShield") {
        Write-Log -Message "$appName uses the InstallShield installer." -Source $appDeployToolkitExtName
        Install-UsingInstallShieldInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "WiXBurn") {
        Write-Log -Message "$appName uses the Wix Burn installer." -Source $appDeployToolkitExtName
        Install-UsingWixBurnInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Wise")  {
        Write-Log -Message "$appName uses the Wise installer." -Source $appDeployToolkitExtName
        Install-UsingWiseInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "InstallAWARE")  {
        Write-Log -Message "$appName uses the InstallAWARE installer." -Source $appDeployToolkitExtName
        Install-UsingInstallAWARE $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Install4j")  {
        Write-Log -Message "$appName uses the Install4j installer." -Source $appDeployToolkitExtName
        Install-UsingInstall4j $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "SetupFactory")  {
        Write-Log -Message "$userDefinedAppName uses the Setup Factory installer." -Source $appDeployToolkitExtName
        Install-UsingSetupFactory $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "Office365ClickToRun")  {
        Write-Log -Message "$appName uses the Microsoft click-to-run installer." -Source $appDeployToolkitExtName
        Install-UsingOffice365ClickToRunInstaller $defaultExeFile $deploymentType
    }

    If ($installerTechnology -eq "WindowsUpdateStandaloneInstaller")  {
        Write-Log -Message "$appName uses the Microsoft Windows update standalone installer." -Source $appDeployToolkitExtName
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
        
    ## Load the first 100000 characters of Unicode for the .exe and search it for installer technology reference
    if (-not ([string]::IsNullOrEmpty($defaultExeFile))) {
        $contentUTF8 = Get-Content -Path $defaultExeFile -Encoding UTF8 -TotalCount 100000
        $contentUnicode = Get-Content -Path $defaultExeFile -Encoding Unicode -TotalCount 100000
    }

    if ((($contentUTF8 -match "NSIS") -or ($contentUnicode -match "NSIS") -or ($contentUTF8 -match "Nullsoft") -or ($contentUnicode -match "Nullsoft")) -and ( -not ($contentUTF8 -match "ClickToRun")) -and ( -not ($contentUnicode -match "ClickToRun")) -and ( -not ($contentUTF8 -match "Inno Setup")) -and ( -not ($contentUnicode -match "Inno Setup")) -and ( -not ($contentUTF8 -match "InstallShield")) -and ( -not ($contentUnicode -match "InstallShield")))  {
        $installerTechnology = "NSIS"
    }

    elseif ((($contentUTF8 -match "Inno Setup") -or ($contentUnicode -match "Inno Setup")) -and ( -not ($contentUTF8 -match "InstallAWARE")) -and ( -not ($contentUnicode -match "InstallAWARE")))  {
        $installerTechnology = "Inno Setup"
    }

    elseif (($contentUTF8 -match "InstallShield") -or ($contentUnicode -match "InstallShield")) {
        $installerTechnology = "InstallShield"
    }

    elseif (($contentUTF8 -match "wixburn") -or ($contentUnicode -match "wixburn")) {
        $installerTechnology = "WiXBurn"
    }

    elseif (($contentUTF8 -match "WiseMain") -or ($contentUnicode -match "WiseMain") -or ($contentUTF8 -match "Wise Installation") -or ($contentUnicode -match "Wise Installation"))  {
        $installerTechnology = "Wise"
    }

    elseif (($contentUTF8 -match "InstallAWARE") -or ($contentUnicode -match "InstallAWARE"))  {
        $installerTechnology = "InstallAWARE"
    }

    elseif (($contentUTF8 -match "install4j") -or ($contentUnicode -match "install4j"))  {
        $installerTechnology = "Install4j"
    }

    elseif (($contentUTF8 -match "Setup Factory") -or ($contentUnicode -match "Setup Factory"))  {
        $installerTechnology = "SetupFactory"
    }

    elseif ($contentUTF8 -match "ClickToRun")  {
        $installerTechnology = "Office365ClickToRun"
    }

    elseif (-not ([string]::IsNullOrEmpty($defaultMsuFile))) {
        $installerTechnology = "WindowsUpdateStandaloneInstaller"
    }

    else {
        ## If the installer technology wasn't detected then log that and return to Deploy-Application to execute user defined installation tasks.
        ## No log will be written if the logging is disabled, which is used in the prevent multiple unnecessary logs being written while 
        ## name, vendor and version information is being gathered for some installer types.
        Write-Log -Message "Installer technology for $appName has not been detected." -Source $appDeployToolkitExtName

        Return
    } 

    return $installerTechnology
}

Function Find-UninstallStringInRegsitry {
<#
.SYNOPSIS
	Detect the uninstall command for an application in the Windows registry.
.DESCRIPTION
    Search through the HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall key in the registry to find the uninstall command for an application.
    On 32bit systems only this key is searched, on 64bit systems both this key and the WOW6432Node key are searched.
    The results are logged so if no uninstall string is found the user can modify the uninstall behaviour accordingly.
.EXAMPLE
	Find-UninstallStringInRegsitry
.NOTES
	This is an internal script function and should typically not be called directly.
#>

    param (
    )

    [string]$uninstallString = ""
    [string]$quietUninstallString = ""
    [hashtable]$returnValues = @{}
    
    ## Attempt to find an uninstall string for the app in the Windows uninstall registry key
    ## Start by checking HKEY_LOCAL_MACHINE, both 32-bit and 64-bit on 64-bit systems (only 32-bit on 32-bit systems)
    $uninstallKey = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    if (((gwmi Win32_OperatingSystem).OSArchitecture) -eq "64-bit") {
        $uninstallKey6432 = Get-ChildItem 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    foreach ($key in $uninstallKey) {
        $registryKey = ($key.Name) -replace "HKEY_LOCAL_MACHINE", "HKLM:"
        $registryProperties = Get-ItemProperty -Path $registryKey
        
        if (($registryKey -like "*$appName*") -or (($registryProperties.DisplayName) -like "*$appName*")) {
            if (-not ([string]::IsNullOrEmpty($registryProperties.UninstallString))) {
                [string]$uninstallString = $registryProperties.UninstallString
            }
            
            if (-not ([string]::IsNullOrEmpty($registryProperties.QuietUninstallString))) {
                [string]$quietUninstallString = $registryProperties.QuietUninstallString
            }
        }
    }
    
    ## If no uninstall string was found in the 64-bit uninstall key, look in the 32-bit uninstall key
    if (-not ([string]::IsNullOrEmpty($uninstallKey6432)) -and ([string]::IsNullOrEmpty($uninstallString)) -and ([string]::IsNullOrEmpty($quietUninstallString))) {
        foreach ($key in $uninstallKey6432) {
            $registryKey = ($key.Name) -replace "HKEY_LOCAL_MACHINE", "HKLM:"
            $registryProperties = Get-ItemProperty -Path $registryKey
         
            if (($registryKey -like "*$appName*") -or (($registryProperties.DisplayName) -like "*$appName*")) {
                if (-not ([string]::IsNullOrEmpty($registryProperties.UninstallString))) {
                    [string]$uninstallString = $registryProperties.UninstallString
                }
            
                if (-not ([string]::IsNullOrEmpty($registryProperties.QuietUninstallString))) {
                    [string]$quietUninstallString = $registryProperties.QuietUninstallString
                }
            }
        }
    }
    
    ## If nothing is found in HKLM, check the current users HKEY_CURRENT_USER uninstall registry key. If the script is running as SYSTEM (as is possible if 
    ## the script has been called by SCCM) then skip the HKCU section.
    if (([Security.Principal.WindowsIdentity]::GetCurrent().Name) -ne "NT AUTHORITY\SYSTEM") {
        if (([string]::IsNullOrEmpty($uninstallString)) -and ([string]::IsNullOrEmpty($quietUninstallString))) {
            $uninstallKey = Get-ChildItem 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    
            foreach ($key in $uninstallKey) {
                $registryKey = ($key.Name) -replace "HKEY_CURRENT_USER", "HKCU:"
                $registryProperties = Get-ItemProperty -Path $registryKey
        
                if (($registryKey -like "*$appName*") -or (($registryProperties.DisplayName) -like "*$appName*")) {
                    if (-not ([string]::IsNullOrEmpty($registryProperties.UninstallString))) {
                        [string]$uninstallString = $registryProperties.UninstallString
                    }
            
                    if (-not ([string]::IsNullOrEmpty($registryProperties.QuietUninstallString))) {
                        [string]$quietUninstallString = $registryProperties.QuietUninstallString
                    }
                }
            }
        }
    }
    
    ## Return the results in a hashtable so that they can be queried by the requesting uninstall function
    if (-not ([string]::IsNullOrEmpty($quietUninstallString))) {
        $uninstallExe, $arguments = $quietUninstallString -split ' +(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)'
        $returnValues.uninstallExe = $uninstallExe.Trim()
        $returnValues.arguments = $arguments.Trim()
    }   
    
    elseif (-not ([string]::IsNullOrEmpty($uninstallString))) {
        $uninstallExe, $arguments = $uninstallString -split ' +(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)'
        $returnValues.uninstallExe = $uninstallExe.Trim()
        $returnValues.arguments = $arguments.Trim()
    }
    
    else {
        $returnValues.uninstallExe = ""
        $returnValues.arguments = ""    
    }
    
    return $returnValues
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
        $arguments = "/AllUsers /S"
    }

    if ($deploymentType -eq "Uninstall") {
        $uninstallValues = Find-UninstallStringInRegsitry
        
        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
        }
        
        elseif (-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = "/AllUsers /S"    
        }
        
        else {
            Write-Log -Message "$appName uninstallation string could not be found in the registry. Uninstall will not proceed." -Source $appDeployToolkitExtName
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
            $arguments = '/sp- /verysilent /norestart /SuppressMSGBoxes /LoadInf=' + '"' + $infFile + '"' + ' /LOG=' + '"' + "$configToolkitLogDir\$appExeLogFileName" + '_Install.log' + '"'
        }
    }

    if ($deploymentType -eq "Uninstall") {
        ## Inno Setup sometimes places an uninstaller in the install directory of the app, attempt to find this in the registry and then uninstall.
        $uninstallValues = Find-UninstallStringInRegsitry
        
        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
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
        $uninstallValues = Find-UninstallStringInRegsitry

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
            Write-Log -Message "$appName uninstallation string could not be found in the registry. Uninstall will not proceed." -Source $appDeployToolkitExtName
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
        $uninstallValues = Find-UninstallStringInRegsitry

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
            Write-Log -Message "$appName uninstallation string could not be found in the registry. Uninstall will not proceed." -Source $appDeployToolkitExtName
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
        $uninstallValues = Find-UninstallStringInRegsitry

        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments + " /S"
        }
        
        else {
            Write-Log -Message "$appName uninstallation string could not be found in the registry. Uninstall will not proceed." -Source $appDeployToolkitExtName
            Return
        }
    }

    ## Install or uninstall string has been worked out, execution process
    ## Log file is written to %TEMP% (no way to specify alternative location with Setup Factory)
    ## Log file will be copied to user defined log location when the install / uninstall finishes.

    $installationStartTime = Get-Date
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
        $uninstallValues = Find-UninstallStringInRegsitry
        
        if ((-not ([string]::IsNullOrEmpty($uninstallValues.uninstallExe))) -and (-not ([string]::IsNullOrEmpty($uninstallValues.arguments)))) {
            $defaultExeFile = $uninstallValues.uninstallExe
            $arguments = $uninstallValues.arguments
        }

        else {
            Write-Log -Message "$appName uninstallation string could not be found in the registry. Uninstall will not proceed." -Source $appDeployToolkitExtName
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

##endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
##region ScriptBody

## Figure out whether this script has been invoked by the Powershell App Deployment Toolkit.
If ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $appDeployToolkitExtName
}
Else {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] invoked directly" -Source $appDeployToolkitExtName
}

## Find which installer technology is being used by the installer.
Write-Log -Message "Attempting to find which installer technology $appName is using." -Source $appDeployToolkitExtName
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
            [string]$appName = $fileDescription
        }
    }

    ## Office 365 click-to-run version is not held in the setup.exe file, but in the folder name Office\Data\[version].
    ## Purposefully leave out AppVendor as Microsoft uses their name in the app name (so it reads Microsoft Coproration Microsoft Office 2016...).
    elseif ($installerTechnology -eq "Office365ClickToRun") {
        [string]$appName =  (Get-Item -Path $defaultExeFile).VersionInfo.FileDescription.Trim()
        [string]$appVersion = (Get-ChildItem -Path "$dirFiles\Office\Data\" -Recurse | ?{ $_.PSIsContainer }).Name
    }

    ## If none of the conditions above are met, use the details in the .exe to fill in $appVendor, $appName and $appVersion.
    else {

        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.CompanyName))) {
            [string]$appVendor = (Get-Item -Path $defaultExeFile).VersionInfo.CompanyName.Trim()
        }
        
        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.ProductName))) {       
            [string]$appName =  (Get-Item -Path $defaultExeFile).VersionInfo.ProductName.Trim()
        }

        if (-not ([string]::IsNullOrEmpty((Get-Item -Path $defaultExeFile).VersionInfo.ProductVersion))) {
            [string]$appVersion = (Get-Item -Path $defaultExeFile).VersionInfo.ProductVersion.Trim()
        }
    }

    ## Remove any special characters in the app name for use in the log name. Also remove the words "installer", "install", "installation or "setup" if they are found in the app name.
    [string]$appName = ($appName -replace "installer", "").Trim()
    [string]$appName = ($appName -replace "install", "").Trim()
    [string]$appName = ($appName -replace "installation", "").Trim()
    [string]$appName = ($appName -replace "setup", "").Trim()
    [string]$appExeLogFileName = ($appName -replace '[\W]', '')
    Write-Log -Message "App Vendor [$appVendor]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Name [$appName]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Version [$appVersion]." -Source $appDeployToolkitExtName
}

## Msu installers don't contain update information in the file properties, so just use Microsoft as manufacturer and the KB number as the name.
if (-not ([string]::IsNullOrEmpty($defaultMsuFile))) {
    [string]$appVendor = "Microsoft"
    $defaultMsuFile -match '\b[a-zA-Z]{2}\d*\b' | Out-Null
    [string]$appName =  $Matches[0]
    [string]$appMsuLogFileName = ($appName -replace '[\W]', '')
    Write-Log -Message "App Vendor [$appVendor]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Name [$appName]." -Source $appDeployToolkitExtName
    Write-Log -Message "App Version [$appVersion]." -Source $appDeployToolkitExtName
}

##endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================