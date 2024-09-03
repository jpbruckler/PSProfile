Describe 'Add-Gitkeep' {
    # BeforeAll block for setting up common mocks
    BeforeAll {
        $fnName = 'Add-Gitkeep'
        $verb = $fnName.Split('-')[0]

        $fnDir = (Resolve-Path ([System.IO.Path]::Combine('..', '..', 'PSProfile','src','PSProfile','Functions'))).Path
        . (Resolve-Path (Join-Path -Path $fnDir -ChildPath "$verb\$fnName.ps1")).Path
    }
}