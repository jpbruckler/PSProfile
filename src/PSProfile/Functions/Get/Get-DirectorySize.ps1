function Get-DirectorySize {
    <#
    .SYNOPSIS
        Get the total size of files within a directory.

    .DESCRIPTION
        The Get-DirectorySize function calculates the total size of all files
        within a specified directory and its subdirectories. The result can be
        displayed in bytes, kilobytes, megabytes, or gigabytes.

    .PARAMETER Path
        The path to the directory you want to calculate the size for. Defaults
        to the current directory if not specified.

    .PARAMETER InType
        The unit of measurement for the directory size. Supported values are B
        (bytes), KB (kilobytes), MB (megabytes), and GB (gigabytes). Default is MB.

    .EXAMPLE
        Get-DirectorySize -Path "C:\Temp" -In GB

        This example calculates the total size of all files within the C:\Temp
        directory and its subdirectories, and displays the result in gigabytes.

    .INPUTS
        System.String

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 0)]
        [string] $Path = '.',

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet('B', 'KB', 'MB', 'GB')]
        [string] $In = "MB"
    )

    process {
        $colItems = (Get-ChildItem -Path $Path -Recurse -File | Measure-Object -Property Length -Sum)

        $sizeInBytes = $colItems.Sum
        $formattedSize = switch ($In) {
            "GB" { "{0:N2} GB" -f ($sizeInBytes / 1GB) }
            "MB" { "{0:N2} MB" -f ($sizeInBytes / 1MB) }
            "KB" { "{0:N2} KB" -f ($sizeInBytes / 1KB) }
            "B" { "{0:N2} B" -f $sizeInBytes }
        }

        return $formattedSize
    }
}