<#
.SYNOPSIS
    This file loads all function files in the Profile/functions directory.
.DESCRIPTION
    This file loads all function files in the Profile/functions directory. It is
    loaded by $env:USERPROFILE\Documents\PowerShell\profile.ps1
#>

Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'functions') -Filter '*.ps1' | ForEach-Object {
    . $_.FullName
}