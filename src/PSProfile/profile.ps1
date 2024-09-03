$PSGalleryResource = (Get-PSResourceRepository -Name PSGallery -ErrorAction SilentlyContinue).Trusted
$PSGalleryRepository = (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy

if ($null -ne $PSGalleryResource -and $PSGalleryResource -ne 'Trusted') {
    Set-PSResourceRepository -Name PSGallery -Trusted
}

if ($null -ne $PSGalleryRepository -and $PSGalleryRepository -ne 'Trusted') {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
}