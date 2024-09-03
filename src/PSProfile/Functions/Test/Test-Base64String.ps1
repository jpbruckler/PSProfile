function Test-Base64String {
    <#
    .SYNOPSIS
        Checks if a given string is in Base64 format.

    .DESCRIPTION
        The Test-Base64String function takes an input string and checks if it is
        a valid Base64 encoded string using a regular expression pattern. The
        function returns $true if the string is a valid Base64 encoded string,
        and $false otherwise.

    .PARAMETER InputString
        The input string to be checked for Base64 format. This parameter is
        mandatory and accepts pipeline input.

    .EXAMPLE
        $isBase64 = Test-Base64String -InputString "SGVsbG8sIHdvcmxkIQ=="

        This example checks if the input string "SGVsbG8sIHdvcmxkIQ==" is a valid
        Base64 encoded string. The function returns $true in this case.

    .INPUTS
        System.String

    .OUTPUTS
        System.Boolean
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [string] $InputString
    )

    process {
        $base64Pattern = '^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$'
        return $InputString -match $base64Pattern
    }
}