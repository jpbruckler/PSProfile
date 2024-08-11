function ConvertTo-MarkdownTable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Object[]] $InputObject
    )

    begin {
        # Create a StringBuilder object
        $sb = New-Object System.Text.StringBuilder
    }

    process {
        if ($null -eq $InputObject) {
            throw "InputObject cannot be null"
        }

        if (-not $properties) {
            $properties = $InputObject | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
            if ($properties.Count -eq 0) {
                throw "No properties found for the given object."
            }

            # Create header row
            $header = $properties -join " | "
            $headerRow = "| $header |"

            # Create separator row
            $separatorRow = "| " + ("--- | " * $properties.Count)

            # Append the header and separator to StringBuilder
            [void]$sb.AppendLine($headerRow)
            [void]$sb.AppendLine($separatorRow)
        }

        # Process each object and append its row
        foreach ($obj in $InputObject) {
            $row = $properties | ForEach-Object { $obj.$_ }
            $row = $row -join " | "
            $row = "| $row |"
            [void]$sb.AppendLine($row)
        }
    }

    end {
        # Return the markdown table
        return $sb.ToString()
    }
}