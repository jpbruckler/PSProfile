function ConvertFrom-ADFileTime {
    <#
    .SYNOPSIS
        Converts an Active Directory FileTime value to a DateTime object.

    .DESCRIPTION
        The ConvertFrom-ADFileTime function takes an Active Directory FileTime value as input,
        which represents the number of 100-nanosecond intervals since January 1, 1601 (UTC),
        and converts it to a DateTime object.

    .PARAMETER ADFileTime
        The Active Directory FileTime value to convert to a DateTime object.

    .EXAMPLE
        $pwdLastSet = 133251702158073284
        $dateTime = ConvertFrom-ADFileTime -ADFileTime $pwdLastSet
        Write-Host $dateTime

        Converts the Active Directory FileTime value 133251702158073284 to a DateTime object and displays it.

    .EXAMPLE
        $pwdLastSet = 133251702158073284
        $dateTime = $pwdLastSet | ConvertFrom-ADFileTime
        Write-Host $dateTime

        Demonstrates how to use pipeline input with ConvertFrom-ADFileTime.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [int64]$ADFileTime
    )

    process {
        $dateTime = [DateTime]::FromFileTimeUtc($ADFileTime).ToLocalTime()
        return $dateTime
    }
}