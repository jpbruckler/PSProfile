function ConvertTo-Base64String {
    <#
    .SYNOPSIS
        Converts a given string to a Base64 encoded string.

    .DESCRIPTION
        The ConvertTo-Base64String function takes an input string and converts
        it to a Base64 encoded string. The function uses UTF-8 encoding for the
        conversion.

    .PARAMETER String
        The input string to be converted to a Base64 encoded string. This
        parameter is mandatory and accepts pipeline input.

    .EXAMPLE
        $base64String = ConvertTo-Base64String -String "Hello, world!"

        This example converts the input string "Hello, world!" to a Base64 encoded string.

    .INPUTS
        System.String

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [string] $String
    )

    process {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
        $base64 = [System.Convert]::ToBase64String($bytes)
        Write-Output $base64
    }
}