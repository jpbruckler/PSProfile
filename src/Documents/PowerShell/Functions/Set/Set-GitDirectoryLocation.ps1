function Set-GitDirectoryLocation {
    <#
    .SYNOPSIS
        Changes the current location to a directory within the local Git folder.
    .DESCRIPTION
        Changes the current location to a directory within the local Git folder.
        Depends on the $env:LOCAL_GIT_FOLDER variable set in the current session.
    .PARAMETER Path
        The path to the directory within the local Git folder.
    .EXAMPLE
        Set-GitDirectoryLocation -Path 'MyProject'

        Changes the current location to $env:LOCAL_GIT_FOLDER\MyProject.
    .EXAMPLE
        Set-GitDirectoryLocation -Path 'MyProject\src'

        Changes the current location to $env:LOCAL_GIT_FOLDER\MyProject\src.
    #>
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Path
    )

    $Location = Join-Path -Path $env:LOCAL_GIT_FOLDER -ChildPath $Path
    
    # Set-Location already handles non-existent paths, so we don't need to check
    Set-Location -Path $Location
}