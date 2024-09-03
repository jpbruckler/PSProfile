function Export-PSResource {
    <#
    .SYNOPSIS
        Downloads a PowerShell module from the PowerShell Gallery as a NuGet package.
    .DESCRIPTION
        The Export-PSResource function downloads a PowerShell module from the
        PowerShell Gallery as a NuGet package. Used for offline installations on
        remote systems, or to add modules to a private repository.
    .PARAMETER ResourceName
        The name of the module to download. This parameter is mandatory and can
        accept input from the pipeline.
    .PARAMETER OutputPath
        The path to save the downloaded NuGet package. The default path is the
        Downloads folder in the user's profile.
    .PARAMETER Force
        If the OutputPath does not exist, the function will create it. Use this
        switch to force the creation of the path.
    #>
    [CmdletBinding()]
    param (
        [Parameter( Mandatory = $true,
            ValueFromPipeline = $true )]
        [String[]] $ResourceName,
        [System.IO.FileInfo] $OutputPath = (Join-Path -Path $env:USERPROFILE -ChildPath 'Downloads'),
        [switch] $Force
    )

    begin {
        if (-not (Test-Path $OutputPath)) {
            if (-not ($Force)) {
                Write-Error ('{0} does not exist. Use -Force to create it.' -f $OutputPath)
                return
            }
            else {
                New-Item -Path $OutputPath -ItemType Directory | Out-Null
            }
        }
    }

    process {
        foreach ($Name in $ResourceName) {
            $Resource = Find-PSResource -Name $Name
            if ([string]::IsNullOrEmpty($Resource)) {
                Write-Error "Resource $Name not found"
                return
            }

            if ([string]::IsNullOrEmpty($Resource.RepositorySourceLocation)) {
                $Source = 'https://www.powershellgallery.com/api/v2'
                Write-Warning -Message "No repository source location found for $Name. Using PSGallery."
            }
            else {
                $Source = $Resource.RepositorySourceLocation
            }

            $URI = '{0}/package/{1}/{2}' -f $Source, $Resource.Name, $Resource.Version
            $OutFile = (Join-Path $OutputPath -ChildPath ('{0}-{1}.nupkg' -f $Resource.Name, $Resource.Version))
            Write-Host "Downloading $URI to $OutFile"
            Invoke-WebRequest -Uri $URI -OutFile $OutFile -ErrorAction SilentlyContinue
        }
    }
}