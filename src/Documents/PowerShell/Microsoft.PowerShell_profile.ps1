# Requires -Version 7

# --------------------------------------
# Current user, current host profile
# --------------------------------------
$env:PSPROFILE_DEBUG = $false
$src = Join-Path -Path ($PROFILE | Split-Path -Parent) -ChildPath 'Loaders'
$perf = @()
$loaders = @(
    'env.ps1',
    'aliases.ps1',
    'completions.ps1',
    'functions.ps1',
    'options.ps1',
    'prompts.ps1',
    'psdrives.ps1',
    'modules.ps1',
    'psreadline.ps1'
)

foreach ($loader in $loaders) {
    $loaderPath = Join-Path -Path $src -ChildPath $loader
    if (Test-Path -Path $loaderPath) {
        Write-Verbose -Message "Importing loader: $loader"
        $perf += @{
            Loader = $loader
            LoadTimeMS = (Measure-Command { . $loaderPath }).TotalMilliseconds
        }
    }
}

function Get-FSUri {
    [CmdletBinding()]
    param(
        [Parameter( Mandatory,
                    ValueFromPipeline )]
        [string[]] $Path,
        [switch] $AsMarkdown
    )

    process {
        foreach ($item in $Path) {
            $itemPath = (Resolve-Path $item).Path
            $SystemUri = [system.uri]::new($itemPath)

            if ($AsMarkdown) {
                $linkName = $SystemUri.LocalPath | Split-Path -LeafBase
                $output = "[{0}]({1})" -f $linkName, $SystemUri.AbsoluteUri
            }
            else {
                $output = $SystemUri.AbsoluteUri
            }
            Write-Output $output
        }
    }
}

if ($env:PSPROFILE_DEBUG -eq $true) {
    $TotalMS = ($perf | Measure-Object -Property LoadTimeMS -Sum).Sum
    Write-Debug -Message "Importing PSProfile loaders took ${TotalMS}ms" -Debug
    Write-Debug -Message "Loaders performance:"             -Debug
    Write-Debug -Message "`tLoader             Time"        -Debug
    Write-Debug -Message "`t-----------        ----------"  -Debug

    foreach ($p in $perf) {
        $msg = "`t{0}{1}{2}ms" -f $($p.Loader), (' ' * (20 - $p.Loader.Length)), $($p.LoadTimeMS)
        Write-Debug -Message $msg -Debug
    }
}
