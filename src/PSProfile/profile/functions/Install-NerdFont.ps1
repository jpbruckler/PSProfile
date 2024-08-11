function Install-NerdFont {
    [CmdletBinding(DefaultParameterSetName = 'NerdFontName')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'NerdFontName')]
        [string[]] $NerdFontName,

        [Parameter(Mandatory = $false, ParameterSetName = 'All')]
        [switch] $All
    )

    begin {}

    process {
        $StagingDir = Join-Path $env:TEMP 'NerdFonts'
        if (-not(Test-Path -Path $StagingDir)) {
            $null = New-Item -Path $StagingDir -ItemType Directory -Force
        }
        $dlList     = @()
        $nfRelease  = Invoke-RestMethod 'https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest'
        $nfList     = $nfRelease.assets |
                Where-Object content_type -eq 'application/zip' |
                Select-Object @{n='Name';e={$_.name.split('.')[0]}}, @{n='Url';e={$_.browser_download_url}} -Unique

        if ($PSCmdlet.ParameterSetName -eq 'NerdFontName') {
            # Compile a list of Nerd Fonts to download
            foreach ($nf in $NerdFontName) {
                if (-not($nfList.Name -contains $nf)) {
                    Write-Warning "Nerd Font '$nf' not found."
                } else {
                    $dlList += $nfList | Where-Object Name -eq $nf
                }
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'All') {
            # Download all Nerd Fonts and install them
            $dlList = $nfList
        }

        $total = $dlList.Count
        $count = 0

        $dlList | Foreach-Object {
            $count++
            $percentComplete = [math]::Round(($count / $total) * 100, 0)
            Write-Progress -Activity "Installing Nerd Fonts" -Status "Processing $($_.Name)" -PercentComplete $percentComplete

            Write-Verbose "Downloading $($_.Name)..."

            $nfZip = Join-Path $StagingDir "$($_.Name).zip"
            $nfDir = Join-Path $StagingDir $_.Name

            # Suppress progress output while downloading the Nerd Fonts
            $global:ProgressPreference = 'SilentlyContinue'
            if (-not(Test-Path -Path $nfDir)) {
                Write-Verbose "Downloading $($_.Name)..."

                Invoke-WebRequest -Uri $_.Url -OutFile $nfZip -UseBasicParsing
                Expand-Archive -Path $nfZip -DestinationPath $nfDir
            }
            # Restore the progress output preference
            $global:ProgressPreference = 'Continue'
        }

        # Install the Nerd Fonts
        Write-Progress -Activity "Installing Nerd Fonts" -Status "Finalizing installation..." -PercentComplete 100
        Install-Font -FontPath $StagingDir -Verbose:$VerbosePreference
    }

    end {
        if (Test-Path -Path $StagingDir) {
            $null = Remove-Item -Path $StagingDir -Recurse -Force
        }
    }
}