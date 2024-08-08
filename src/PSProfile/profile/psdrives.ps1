# Description: Create a PSDrive for the local git folder
# $env:LOCAL_GIT_FOLDER set in env.ps1
$gitPSDriveName = 'git'
if (-not (Test-Path -Path $env:LOCAL_GIT_FOLDER)) {
    $null = New-Item -Path $env:LOCAL_GIT_FOLDER -ItemType Directory -Force
}

if ($null -eq (Get-PSDrive | Where-Object { $_.Name -eq $gitPSDriveName })) {
    New-PSDrive -Name $gitPSDriveName -PSProvider FileSystem -Root $env:LOCAL_GIT_FOLDER
}

if ($null -ne $env:STARSHIP_CONFIG) {
    if (-not (Test-Path -Path $env:STARSHIP_CONFIG)) {
        $null = New-Item -Path $env:STARSHIP_CONFIG -ItemType File -Force
    }
}