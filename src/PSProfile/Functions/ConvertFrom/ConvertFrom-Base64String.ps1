function ConvertFrom-Base64String {
    <#
    .SYNOPSIS
        Converts a given Base64 encoded string to its original string representation.

    .DESCRIPTION
        The ConvertFrom-Base64String function takes an input Base64 encoded string
        and converts it back to its original string representation. The function
        uses UTF-8 encoding for the conversion.

    .PARAMETER Base64String
        The input Base64 encoded string to be converted back to its original
        tring representation. This parameter is mandatory and accepts pipeline input.

    .EXAMPLE
        $originalString = ConvertFrom-Base64String -Base64String "SGVsbG8sIHdvcmxkIQ=="

        This example converts the input Base64 encoded string "SGVsbG8sIHdvcmxkIQ=="
        back to its original string representation: "Hello, world!".

    .INPUTS
        System.String

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [string] $Base64String
    )

    process {
        $bytes = [System.Convert]::FromBase64String($Base64String)
        $originalString = [System.Text.Encoding]::UTF8.GetString($bytes)
        Write-Output $originalString
    }
}