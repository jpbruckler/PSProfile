# Starship
Invoke-Expression (& $(Get-Command starship | Select-Object -ExpandProperty Source) init powershell --print-full-init | Out-String)