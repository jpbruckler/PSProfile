function Convert-LocalTimeToUTC {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [datetime]$LocalTime = (Get-Date)
    )

    $utcTime = $LocalTime.ToUniversalTime()
    return $utcTime
}