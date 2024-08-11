function Install-Font {
    <#
    .SYNOPSIS
        Installs a font file or all font files in a directory.
    .DESCRIPTION
        Installs a font file or all font files in a directory. This function uses
        the Windows Shell to install the font. This allows installing fonts without
        needing admin rights, but does result in a pop-up window for each font.
    .PARAMETER FontPath
        The path to the font file or directory containing font files to install.
    .EXAMPLE
        Install-Font -FontPath 'C:\path\to\font.ttf'
        Installs the font file at 'C:\path\to\font.ttf'.
    .EXAMPLE
        Install-Font -FontPath 'C:\path\to\fonts'
        Installs all font files in the directory 'C:\path\to\fonts'.
    .NOTES
        - https://richardspowershellblog.wordpress.com/2008/03/20/special-folders/
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true)]
        [string]$FontPath
    )

    begin {
        # Use the Windows Shell to install the font. This allows installing fonts
        # without needing admin rights.
        $Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
    }

    process {
        # If the given path is a directory, install all fonts in the directory.
        if (Test-Path -Path $FontPath -PathType Container) {
            $Include = ('*.fon','*.otf','*.ttc','*.ttf')
            $fileList = Get-ChildItem -Path $FontPath -Include $Include -Recurse
            $total = $fileList.Length
            $count = 0

            Write-Verbose "Installing fonts in directory: $FontPath"

            $fileList | ForEach-Object {
                $count++
                $percentComplete = [math]::Round(($count / $total) * 100, 0)
                Write-Progress -Activity "Installing fonts..." -Status "Processing $($_.Name)" -PercentComplete $percentComplete
                If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                    Write-Verbose "Installing font: $($_.FullName)"
                    $Destination.CopyHere($_.FullName, 0x10)
                } else {
                    Write-Warning "Font already installed: $($_.FullName)"
                }
            }
        } else {
            If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                Write-Verbose "Installing font: $FontPath"
                $Destination.CopyHere($FontPath, 0x10)
            } else {
                Write-Warning "Font already installed: $FontPath"
            }
        }
    }
}