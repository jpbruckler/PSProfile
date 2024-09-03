<#
.SYNOPSIS
    This file loads all function files in the Profile/functions directory.
.DESCRIPTION
    This file loads all function files in the Profile/functions directory. It is
    loaded by $env:USERPROFILE\Documents\PowerShell\profile.ps1
#>

$fxRoot = Join-Path -Path ($PROFILE | Split-Path -Parent) -ChildPath 'Functions'

Get-ChildItem -Path $fxRoot -Filter '*.ps1' -File -Recurse | ForEach-Object {
    . $_.FullName
}