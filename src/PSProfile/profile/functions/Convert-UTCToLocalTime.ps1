function Convert-UTCToLocalTime {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [datetime]$UTCTime
    )

    $localTime = $UTCTime.ToLocalTime()
    return $localTime
}