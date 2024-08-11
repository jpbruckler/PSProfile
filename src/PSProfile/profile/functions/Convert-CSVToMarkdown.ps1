function Convert-CSVToMarkdown {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $csvData = Import-Csv -Path $Path
    if ($csvData.Count -eq 0) {
        Write-Error "CSV file is empty"
        return
    }

    $headers = $csvData[0].PSObject.Properties.Name
    $separator = $headers | ForEach-Object { '---' }

    # Print the headers
    $headerRow = $headers -join " | "
    $separatorRow = $separator -join " | "

    $markdownTable = @()
    $markdownTable += "| $headerRow |"
    $markdownTable += "| $separatorRow |"

    # Print the rows
    foreach ($row in $csvData) {
        $rowData = @()
        foreach ($header in $headers) {
            $rowData += $row.$header
        }
        $markdownTable += "| " + ($rowData -join " | ") + " |"
    }

    return $markdownTable -join "`n"
}