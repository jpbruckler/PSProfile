function Convert-IntToChar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [int[]] $Int,
        [switch] $AsArray,
        [switch] $AsString
    )

    begin {
        $Output = [System.Collections.Generic.List[char]]::new()
    }

    process {
        foreach ($i in $Int) {
            $null = $Output.Add([char]$i)
        }
    }

    end {
        if ($AsArray) {
            return $Output.ToArray()
        }
        elseif ($AsString) {
            return -join $Output
        }
        else {
            return $Output
        }
    }
}