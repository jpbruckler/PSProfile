function ConvertFrom-UnixTimestamp {
    <#
    .SYNOPSIS
        Converts a Unix timestamp to a DateTime object.

    .DESCRIPTION
        The ConvertFrom-UnixTimestamp function takes a Unix timestamp as input (in seconds or milliseconds)
        and converts it to a DateTime object.

    .PARAMETER UnixTimestamp
        The Unix timestamp to convert to a DateTime object. This can be in seconds or milliseconds.

    .EXAMPLE
        $unixTimestamp = 1682705422
        $dateTime = ConvertFrom-UnixTimestamp -UnixTimestamp $unixTimestamp
        Write-Host $dateTime

        Converts the Unix timestamp 1682705422 to a DateTime object and displays it.

    .EXAMPLE
        $unixTimestamp = 1682705422
        $dateTime = $unixTimestamp | ConvertFrom-UnixTimestamp
        Write-Host $dateTime

        Demonstrates how to use pipeline input with ConvertFrom-UnixTimestamp.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [double]$UnixTimestamp
    )

    process {
        $origin = [DateTime]::new(1970, 1, 1, 0, 0, 0, 0, [DateTimeKind]::Utc)

        # Check if the Unix timestamp is in milliseconds
        if ($UnixTimestamp -gt [DateTime]::MaxValue.Ticks / [TimeSpan]::TicksPerSecond) {
            # Convert Unix timestamp from milliseconds to seconds
            $UnixTimestamp /= 1000
        }

        $convertedDateTime = $origin.AddSeconds($UnixTimestamp).ToLocalTime()

        return $convertedDateTime
    }
}