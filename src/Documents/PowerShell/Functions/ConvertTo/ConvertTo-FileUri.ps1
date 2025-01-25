function ConvertTo-FileUri {
    param(
        [Parameter( Mandatory,
                    ValueFromPipeline,
                    ValueFromPipelineByPropertyName,
                    Position = 0
                    )]
        [System.IO.FileInfo] $Path
    )

    return ([System.Uri]$Path).AbsoluteUri
}