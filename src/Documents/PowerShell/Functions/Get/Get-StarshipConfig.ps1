function Get-StarshipConfig {
    [CmdletBinding(DefaultParameterSetName='List')]
    param(
        [Parameter( Mandatory = $true, ParameterSetName = 'Named' )]
        [string]$Name,

        [Parameter( Mandatory = $false, ParameterSetName = 'List' )]
        [switch]$List
    )

    if ($PSCmdlet.ParameterSetName -eq 'List' -or $List -eq $true) {
        Get-ChildItem -Path $env:USERPROFILE\.config\starship -Filter *.toml -File | 
            Where-Object -Property Name -NE 'starship.toml' |
            Select-Object Name
    }

    if (-not ($Name.EndsWith('toml'))) {
        $Name = "$Name.toml"
    }

    if ($Name -and (Test-Path (Join-Path $env:USERPROFILE\.config\starship -ChildPath $Name))) {
        Get-Item (Join-Path $env:USERPROFILE\.config\starship -ChildPath $Name)
    }
}