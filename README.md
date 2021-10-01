# SCCM SharePoint patching module
Powershell module for sccm and software center patching

This module is a toolset created to aid in patching through SCCM with Powershell.

All fuctions are created to loop through $servers variable.

All the functions that require pscredentials validates that pscredentials are set with TestPSCredentials function and prompt for it if missing.
This can be set with the function pscredentials

Any functions that changes a state or value have validation before running. 
Any function that is only informational are not validated before running.


# Installation

  1) Download this project in .zip
  2) Extract and run Launch.ps1
  3) 
  The module should now be installed and readuy for use. Open Powershell and type Get-Command -Module Sharepointpatching to validate.
  You have to change the server lists to your enviroment and the main script for it to connect with your SCCM server. Read the notes in the script.


# Commands

    SP-ClearSCCMCache       - Clearing SCCM Cache
    SP-CreateCheckpoints    - Creating a checkpoint
    SP-ForceStopServers     - Forced shutdown locally from VM's
    SP-GetApplications      - Getting application information from software center on VM's
    SP-GetCheckpoints       - Getting available checkpoints
    SP-InstallationStatus   - Checking if an application from software center is installed by hash
    SP-RemoveCheckpoints    - Removing checkpoints
    SP-RunSCCMClientAction  - Running SCCM clien actions
    SP-Servers              - Assigning a server list to run rest of the commands toward. Need to be ran first.
    SP-StartServers         - Starting VM's through VMM
    SP-StopServers          - Stopping VM's through VMM
    SP-TestConnection       - Testing ICMP towards VM's
    SP-TriggerInstallation  - Starting installation of an application on VM's
    SP-VMConnect            - Connecting to VM's through RDP
    SP-VMStatus             - Checking VM status (Running or PowerOff) through VMM
    SP-ListSCCMPackages     - Lists available packages in SCCM with aditional information about each package.
