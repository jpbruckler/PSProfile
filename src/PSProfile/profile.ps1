Set-PSResourceRepository -Name PSGallery -Trusted
PowerShellGet\Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

$Src = Join-Path -Path $PSScriptRoot -ChildPath 'Profile'
if (-not(Test-Path -Path $Src)) {
    $null = New-Item -Path $Src -ItemType Directory -Force
}