function Set-StarshipConfig {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'FileInfo')]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [string]$Path
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'FileInfo') {
            Write-Verbose "Setting TargetPath to $File"
            $TargetPath = $File.FullName
        }
        else {
            if (-not (Test-Path $Path -PathType Leaf)) {
                Write-Error "Provided path must point to a file. [$Path]"
            }
            Write-Verbose "Setting TargetPath to $Path"
            $TargetPath = (Resolve-Path -Path $Path).Path
        }

        try {
            $null = New-Item -ItemType HardLink -Path $env:USERPROFILE\.config\starship\starship.toml -Value $TargetPath -Force -ErrorAction Stop
            Write-Information 'Starship configuration updated.'
        }
        catch {
            Write-Error "Unable to change Starship config. Error: $($_.Exception.Message)"
        }
    }
}