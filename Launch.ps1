Begin
{
    # Current location
    $ScriptPath = $MyInvocation.MyCommand.Path
    $ScriptPath = $ScriptPath.Trim(($ScriptPath).Split("\")[-1])
    # SharePointPatching.psm1
    $Module = $ScriptPath + "SharePointPatching.psm1"
    # Initiate_Serverlists.ps1
    $Serverlists = $ScriptPath + "Initiate_Serverlists.ps1"
    #SharePointPatching
    $ModulePath = $Path[0] + "\Sharepointpatching"
    # SharePointPatching module path
    $Path = $env:PSModulePath -split ";"
    $ModulePath = $Path[0] + "\Sharepointpatching"
    # Path to serverlist initiation
    $InitiateServerlists = "$ModulePath\Initiate_Serverlists.ps1"
    # Path to manifest
    $Manifestpath =  $ModulePath  + "\SharePointPatching.psd1"
}
Process
{
    # Create module folder and move module files
    New-Item  $ModulePath -ItemType Directory 
    Copy-Item $Module -Destination $ModulePath -Force
    Copy-Item $Serverlists -Destination $ModulePath -Force
    
    New-ModuleManifest -Path $Manifestpath `
    -Author "Trond Weiseth" -Copyright "(c)2021 Sharepointpatching" `
    -Description "SharePoint patching toolset" -ModuleVersion 1.0 `
    -RootModule .\SharePointPatching.psm1
}
End
{
    &$InitiateServerlists
}
