function ConvertTo-UnixTimestamp {
    <#
    .SYNOPSIS
        Converts a DateTime object to a Unix timestamp.

    .DESCRIPTION
        The ConvertTo-UnixTimestamp function takes a DateTime object as input and converts it to a Unix timestamp
        in seconds.

    .PARAMETER DateTime
        The DateTime object to convert to a Unix timestamp.

    .EXAMPLE
        $dateTime = Get-Date
        $unixTimestamp = ConvertTo-UnixTimestamp -DateTime $dateTime
        Write-Host $unixTimestamp

        Converts the current date and time to a Unix timestamp and displays it.

    .EXAMPLE
        $dateTime = Get-Date
        $unixTimestamp = $dateTime | ConvertTo-UnixTimestamp
        Write-Host $unixTimestamp

        Demonstrates how to use pipeline input with ConvertTo-UnixTimestamp.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [DateTime]$DateTime
    )

    process {
        $origin = [DateTime]::new(1970, 1, 1, 0, 0, 0, 0, [DateTimeKind]::Utc)
        $dateTimeUtc = $DateTime.ToUniversalTime()

        $timeSpan = $dateTimeUtc - $origin
        $unixTimestamp = [int64]$timeSpan.TotalMilliseconds

        return $unixTimestamp
    }
}