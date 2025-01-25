function ConvertFrom-SafelinkUrl {
    <#
    .SYNOPSIS
        Converts a Microsoft SafeLink URL to the original URL.

    .DESCRIPTION
        The ConvertFrom-SafelinkUrl function takes a Microsoft SafeLink URL as
        input and returns the original URL. Microsoft SafeLink URLs are often
        used in emails for security purposes, and this function provides an
        easy way to retrieve the original URL.

    .PARAMETER Url
        A string that represents the SafeLink URL to be converted. This
        parameter is mandatory and can accept input from the pipeline.

    .EXAMPLE
        $originalUrl = ConvertFrom-SafelinkUrl -Url "https://go.microsoft.com/fwlink/?linkid=123456"
        Converts the given SafeLink URL and stores the original URL in the $originalUrl variable.

    .EXAMPLE
        "https://go.microsoft.com/fwlink/?linkid=123456" | ConvertFrom-SafelinkUrl
        Uses the pipeline to pass the SafeLink URL to the ConvertFrom-SafelinkUrl function and outputs the original URL.

    .INPUTS
        System.String
        You can pipe a string that contains the SafeLink URL to ConvertFrom-SafelinkUrl.

    .OUTPUTS
        System.String
        Returns the original URL as a string.

    .NOTES
        This function uses the System.Web assembly and may require appropriate permissions or assemblies to be loaded.

    .LINK
        [System.Web.HttpUtility]::ParseQueryString
    #>
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $Url
    )

    Add-Type -AssemblyName System.Web
    $l = $Url -replace 'https://go.microsoft.com/fwlink/?linkid=', ''
    return ([System.Web.HttpUtility]::ParseQueryString((New-Object -TypeName System.Uri -ArgumentList $l).Query))["url"]

}