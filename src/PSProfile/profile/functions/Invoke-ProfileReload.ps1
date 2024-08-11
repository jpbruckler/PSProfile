function Invoke-ProfileReload {
    <#
    .SYNOPSIS
        Reloads the PowerShell profile.
    .DESCRIPTION
        Reloads the PowerShell profile by dot-sourcing the profile.ps1 file.
    .NOTES
        From about_Profiles:
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

    # Remove the current functions from the session
    $WinPwshProfile = Join-Path -Path $env:USERPROFILE -ChildPath 'Documents\WindowsPowerShell'
    $PwshProfile    = Join-Path -Path $env:USERPROFILE -ChildPath 'Documents\PowerShell'
    $fxnList        = @()

    if (Test-Path -Path $WinPwshProfile) {
        $fxnList += Get-ChildItem -Path (Join-Path -Path $WinPwshProfile -ChildPath 'profile\functions') -Filter '*.ps1'
    }
    if (Test-Path -Path $PwshProfile) {
        $fxnList += Get-ChildItem -Path (Join-Path -Path $PwshProfile -ChildPath 'profile\functions') -Filter '*.ps1'
    }
    $fxnList | ForEach-Object {
        Remove-Item -Path (Join-Path -Path function: -ChildPath $_.BaseName) -Force -ErrorAction SilentlyContinue
        . $_.FullName
    }

    $profiles = @(
        (Join-Path -Path $PwshProfile -ChildPath 'Profile.ps1'),
        (Join-Path -Path $PwshProfile -ChildPath 'Microsoft.PowerShell_profile.ps1')
        (Join-Path -Path $WinPwshProfile -ChildPath 'Profile.ps1'),
        (Join-Path -Path $WinPwshProfile -ChildPath 'Microsoft.PowerShell_profile.ps1')
    )

    $profiles | ForEach-Object {
        if (Test-Path -Path $_) {
            . $_
        }
    }
}