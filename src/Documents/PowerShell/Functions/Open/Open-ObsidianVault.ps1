function Open-ObsidianVault {
    <#
    .SYNOPSIS
        Opens an Obsidian Vault
    .DESCRIPTION
        Opens an Obsidian vault by the vault name or ID using an Obsidian URI.
        Looks for the Obsidian executable in the PATH and the default installation
        location ($env:LOCALAPPDATA\Programs\Obsidian\Obsidian.exe). If the executable
        is not found, an error is written and the function returns.
    .PARAMETER VaultName
        Name or ID of the Vault to Launch.
    .EXAMPLE
        Open-ObsidianVault -VaultName 'My Vault'
    .EXAMPLE
        Open-ObsidianVault -VaultName 2c98ea7c429980bf
    .NOTES
        - https://help.obsidian.md/Extending+Obsidian/Obsidian+URI
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$VaultName
    )

    $ObsidianPath = Get-Command obsidian -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
    if (-not $ObsidianPath) {
        $ObsidianPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Programs\Obsidian\Obsidian.exe'
    }

    if (-not $ObsidianPath) {
        Write-Error 'Obsidian is not installed on this machine.'
        return
    }

    $Vault = [uri]::EscapeDataString($VaultName)
    $URI = "obsidian://open?vault=$Vault"

    Start-Process -FilePath $ObsidianPath -ArgumentList $URI -RedirectStandardOutput ".\NUL"
}