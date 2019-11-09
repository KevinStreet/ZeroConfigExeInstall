### Zero-Config Executable Installation

The Zero-Config Executable (ZeroConfigExe) Installation script is an extension to the PowerShell App Deployment Toolkit designed to make deploying executables as simple as deploying MSI files. The PSADT allows an administrator to simply drop an MSI file into the Files directory and run Deploy-Application.exe file to automatically install the MSI with default silent parameters. With the addition of this extension the administrator is given the same ability with executable installers.

### Features
This extension supports the following installer technologies:
* Windows Installer
* NSIS
* Inno Setup
* InstallShield
* WiX Burn
* Wise
* InstallAWARE
* install4j
* Setup Factory

It also has special logic for installing the following products:
* Microsoft Office 365 click-to-run
* Microsoft Windows update files (.msu)

The script scans the installer file for a reference to one of the installer technologies and then proceeds to silently install it. Uninstalls are also supported and both installs and uninstalls are logged in the administrator defined logging location when supported by the installer.

This extension furthers the goal of the PowerShell App Deployment Toolkit by offering an easy to use, powerful and consistent deployment experience for both the administrator and end user.

## License

Zero-Config Executable Installation.
Copyright (C) 2019 - Kevin Street.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. 

You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
