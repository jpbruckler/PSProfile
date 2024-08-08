if (-not((Get-PSResourceRepository -Name PSGallery).Trusted)) {
    Set-PSResourceRepository -Name PSGallery -Trusted
}

if ((PowerShellGet\Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
    PowerShellGet\Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
}

$src = Join-Path -Path $PSScriptRoot -ChildPath 'Profile'
if (-not(Test-Path -Path $src)) {
    $null = New-Item -Path $src -ItemType Directory -Force
}

. (Join-Path -Path $src -ChildPath 'aliases.ps1')
. (Join-Path -Path $src -ChildPath 'completions.ps1')
. (Join-Path -Path $src -ChildPath 'env.ps1')
. (Join-Path -Path $src -ChildPath 'functions.ps1')
. (Join-Path -Path $src -ChildPath 'options.ps1')
. (Join-Path -Path $src -ChildPath 'prompts.ps1')
. (Join-Path -Path $src -ChildPath 'psdrives.ps1')