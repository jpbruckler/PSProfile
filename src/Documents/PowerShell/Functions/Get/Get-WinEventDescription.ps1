function Get-WinEventIDDescription {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [int[]] $EventId,

        [Parameter(Mandatory = $false)]
        [string] $ProviderName,

        [Parameter(Mandatory = $false)]
        [string] $LiteralProviderName,

        [string] $Delimiter
    )

    process {
        if ((-not $ProviderName) -and (-not $LiteralProviderName)) {
            $ProviderName = '*'
            Write-Warning 'No provider name specified. Using wildcard. This may take a while.' -WarningAction Continue
        }
        elseif ($ProviderName) {
            $ProviderName = '*{0}*' -f $ProviderName
        }
        elseif ($LiteralProviderName) {
            $ProviderName = $LiteralProviderName
        }
        $Events = (Get-WinEvent -ListProvider $ProviderName).Events | Where-Object Id -In $EventId
        foreach ($Event in $Events) {
            $Description = (($Event | Select-Object -ExpandProperty Description) -split "`r`n")[0].Trim()
            if ([string]::IsNullOrEmpty($Description)) {
                return [PSCustomObject]@{
                    EventID     = $Event.Id
                    Description = $Description
                }
            }
            else {
                return ('{0}{1}{2}' -f $Event.Id, $Delimiter, $Description)
            }
        }
    }
}