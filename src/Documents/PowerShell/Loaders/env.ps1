$env:LOCAL_GIT_FOLDER = Join-Path -Path $env:USERPROFILE -ChildPath 'Documents\git'
$env:STARSHIP_CONFIG = Join-Path -Path $env:USERPROFILE -ChildPath '.config\starship\starship.toml'
if (Test-Path (Join-Path $env:APPDATA 'espanso')) {
    $env:ESPANSO_PATH = (Join-Path $env:APPDATA 'espanso')
}