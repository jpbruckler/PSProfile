function Split-Array {
    <#
    .SYNOPSIS
    Splits an input array into smaller subarrays with a specified maximum number of elements.

    .DESCRIPTION
    The Split-Array function takes an input array and splits it into smaller subarrays, each containing a specified maximum number of elements. The function is useful when dealing with large arrays that need to be processed in smaller chunks.

    .PARAMETER InputArray
    The input array that you want to split into smaller subarrays. This parameter is mandatory and accepts pipeline input.

    .PARAMETER MaximumElements
    The maximum number of elements allowed in each smaller subarray. The default value is 10.

    .EXAMPLE
    $myArray = 1..20
    $splitArrays = Split-Array -InputArray $myArray -MaximumElements 5

    This example splits the input array $myArray (which contains numbers 1 to 20) into smaller subarrays of 5 elements each:

    (1, 2, 3, 4, 5),
    (6, 7, 8, 9, 10),
    (11, 12, 13, 14, 15),
    (16, 17, 18, 19, 20)

    .INPUTS
    System.Object[]

    .OUTPUTS
    System.Collections.ArrayList
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [object[]] $InputArray,

        [Alias('Max', 'Size')]
        [int] $MaximumElements = 10
    )

    begin {
        $OutputArrayList = [System.Collections.ArrayList]::new()
    }

    process {
        for ($i = 0; $i -lt $InputArray.Length; $i += $MaximumElements) {
            $slice = $InputArray[$i..($i + $MaximumElements - 1)]
            $null = $OutputArrayList.Add($slice)
        }
    }

    end {
        Write-Output $OutputArrayList
    }
}