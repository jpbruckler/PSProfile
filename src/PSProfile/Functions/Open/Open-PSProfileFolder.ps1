enum ProfileScope {
    AllUsersAllHosts
    AllUsersCurrentHost
    CurrentUserAllHosts
    CurrentUserCurrentHost
}
function Open-PSProfileFolder {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER Scope
    Parameter description

    .EXAMPLE
    An example

    .NOTES

    All Users, All Hosts
        Windows - $PSHOME\Profile.ps1
        Linux - /opt/microsoft/powershell/7/profile.ps1
        macOS - /usr/local/microsoft/powershell/7/profile.ps1
    All Users, Current Host
        Windows - $PSHOME\Microsoft.PowerShell_profile.ps1
        Linux - /opt/microsoft/powershell/7/Microsoft.PowerShell_profile.ps1
        macOS - /usr/local/microsoft/powershell/7/Microsoft.PowerShell_profile.ps1
    Current User, All Hosts
        Windows - $HOME\Documents\PowerShell\Profile.ps1
        Linux - ~/.config/powershell/profile.ps1
        macOS - ~/.config/powershell/profile.ps1
    Current user, Current Host
        Windows - $HOME\Documents\PowerShell\Microsoft.PowerShell_profile.ps1
        Linux - ~/.config/powershell/Microsoft.PowerShell_profile.ps1
        macOS - ~/.config/powershell/Microsoft.PowerShell_profile.ps1

    #>
    param(
        [ProfileScope] $Scope
    )
}