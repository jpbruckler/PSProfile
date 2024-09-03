function Add-Gitkeep {
    <#
    .SYNOPSIS
        Creates .gitkeep files in empty directories within a Git repository.

    .DESCRIPTION
        The Add-GitDeleteFiles function creates .gitkeep files in empty directories within a Git repository.
        The -Cleanup switch removes .gitkeep files from directories that are no longer empty.

    .PARAMETER Path
        The path to the root directory of the Git repository. Defaults to the current location.

    .PARAMETER Cleanup
        If specified, the function will remove .gitkeep files from directories that are no longer empty.

    .EXAMPLE
        Add-GitDeleteFiles -Path "C:\MyGitRepo"

        Creates .gitkeep files in all empty directories within the specified Git repository.

    .EXAMPLE
        Add-GitDeleteFiles -Path "C:\MyGitRepo" -Cleanup

        Creates .gitkeep files in empty directories and removes .gitkeep files from directories that are no longer empty.

    .EXAMPLE
        Get-ChildItem -Path "C:\MyGitRepos" -Directory | Add-GitDeleteFiles -Cleanup

        Demonstrates how to use pipeline input with Add-GitDeleteFiles to process multiple repositories.
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$Path = (Get-Location).Path,

        [switch]$Cleanup
    )

    process {
        foreach ($dirPath in $Path) {
            $Directories = Get-ChildItem -Path $dirPath -Directory -Recurse -ErrorAction SilentlyContinue

            foreach ($directory in $Directories) {
                $ChildItems = Get-ChildItem -Path $directory -ErrorAction SilentlyContinue
                $gitDeletePath = Join-Path $directory.FullName '.gitkeep'

                if ($ChildItems.Count -eq 0) {
                    $null = New-Item -Path $gitDeletePath -ItemType File -Verbose
                    Write-Verbose "Created .gitkeep at $($directory.FullName)"
                }
                elseif ($Cleanup -and ($ChildItems.Count -gt 1) -and ($ChildItems.Name -contains '.gitkeep')) {
                    $null = Remove-Item -Path $gitDeletePath -Force -Verbose
                    Write-Verbose "Removed .gitkeep from $($directory.FullName)"
                }
            }
        }
    }
}
