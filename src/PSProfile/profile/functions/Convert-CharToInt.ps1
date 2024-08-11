function Convert-CharToInt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [char[]] $Char
    )

    begin {
        $Output = [System.Collections.Generic.List[int]]::new()
    }

    process {
        foreach ($c in $Char) {
            $null = $Output.Add([int]$c)
        }
    }

    end {
        return $Output
    }
}