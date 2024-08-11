# Set environment variables
Set-Item -Path "env:PSProfileGistID" -Value '1ba47a193046115d3d14d28ee2c58f2f'
Set-Item -Path "env:OBSIDIAN_VAULT" -Value '$DriveName:\arcane-algorithms'

# Profile variables
$script:documents = (Join-Path $env:USERPROFILE 'Documents')



# Module Imports
$RequiredModules = @('Terminal-Icons', 'posh-git')
foreach ($module in $RequiredModules) {
    try {
        Import-Module $module -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to import module '$module'. Attempting to install."
        Install-PSResource Terminal-Icons -Scope CurrentUser -TrustRepository
        Import-Module $module
    }
}

function Set-LocationGit {
    param(
        [string] $Path = $null
    )
    Set-Location (Join-Path git:\ $Path)
}
Set-Alias -Name cdg -Value Set-LocationGit


# Update Path
$PathExt = @(
    (Join-Path $env:APPDATA 'espanso'),
    (Join-Path $env:APPDATA 'npm'),
    (Join-Path $env:SystemDrive 'tools'),
    (Join-Path $env:SystemDrive 'tools\pstools'),
    (Join-Path ${env:ProgramFiles(x86)} 'VMWare\VMware Workstation')
)

foreach ($Path in $PathExt) {
    if (Test-Path $Path) {
        $env:PATH = '{0};{1}' -f $env:PATH, $Path
    }
}


#region Git setup
# Using $env:USERPROFILE prevents putting git repos into a OneDrive directory.
# If you _want_ to keep your git repos in OneDrive, change to $env:OneDrive or
# $env:OneDriveCommercial
$script:gitdir = (Join-Path $env:USERPROFILE 'Documents\git')
$script:DriveName = 'GIT'

if (-not (Test-Path $script:gitdir)) {
    try {
        $null = New-Item -ItemType Directory -Path $script:gitdir -Force -ErrorAction stop
    }
    catch {
        Write-Error -Message "Unable to create missing path '$script:gitdir'. Exception: $_"
    }
}

$mapped = Get-PSDrive -Name $script:DriveName -ErrorAction SilentlyContinue

# if $DriveName is mapped, but the root is different then the provided path, unmap
if ($null -ne $mapped -and $mapped.Root -ne $script:gitdir) {
    $null = Remove-PSDrive -Name $script:DriveName -ErrorAction SilentlyContinue
    $mapped = $null
}
if ($null -eq $mapped) {
    New-PSDrive -name $script:DriveName -PSProvider FileSystem -Root $script:gitdir
}

#endregion

#region OhMyPosh setup
#oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\material.omp.json" | Invoke-Expression
#oh-my-posh init pwsh --config "$env:POSH_THEMES_PATH\amro.omp.json" | Invoke-Expression
Invoke-Expression (&starship init powershell)
#endregion

#region ALIASES
Set-Alias -Name 'obsidian' -Value 'C:\Users\john\AppData\Local\Obsidian\Obsidian.exe'
Set-Alias -Name '$DriveName' -Value "C:\Program Files (x86)\GitHub CLI\$DriveName.exe"
#endregion

#region PSREADLINE
Set-PSReadLineOption -HistorySearchCursorMovesToEnd
Set-PSReadlineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadlineKeyHandler -Key DownArrow -Function HistorySearchForward
Set-PSReadlineOption -BellStyle None #Disable ding on typing error
Set-PSReadlineOption -EditMode Emacs #Make TAB key show parameter options
Set-PSReadlineKeyHandler -Key Tab -Function MenuComplete

# Ctrl+e to open Edge with Google
Set-PSReadlineKeyHandler -Key Ctrl+e -ScriptBlock { Start-Process "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe" -ArgumentList "https://www.google.com" }

# Ctrl+Shift+O to open Obsidian
Set-PSReadLineKeyHandler -Chord Ctrl+O -ScriptBlock {
    $vault = (Get-Item -Path $env:OBSIDIAN_VAULT).FullName
    $tf = New-TemporaryFile
    Set-Location $vault -Verbose
    Write-Information "Pulling latest changes from remote" -InformationAction Continue
    Start-Process git -ArgumentList pull -RedirectStandardOutput $tf -Wait -NoNewWindow
    Write-Information "Pull complete, launching obsidian." -InformationAction Continue
    Start-Process 'C:\Users\u55398\AppData\Local\Obsidian\Obsidian.exe' -RedirectStandardOutput $tf -
    Start-Sleep -Seconds 3
    Remove-Item $tf
    return
}

# Ctrl+Shift+C to open VS Code in current directory
Set-PSReadLineKeyHandler -Chord Ctrl+C -ScriptBlock {
    Start-Process code -ArgumentList '.' -NoNewWindow
}

# Ctrl+g to commit and push
Set-PSReadlineKeyHandler -Chord Ctrl+G -ScriptBlock {
    $message = Read-Host "Please enter a commit message"
    git commit -m "$message" | Write-Host
    $branch = (git rev-parse --abbrev-ref HEAD)
    Write-Host "Pushing ${branch} to remote"
    git push origin $branch
}
#endregion PSREADLINE

#region Functions
function Get-VmwareVMStatus {
    <#
    .SYNOPSIS
        Get the status of VMware virtual machines.

    .DESCRIPTION
        The Get-VmwareVMStatus function retrieves the status of VMware virtual machines
        and displays the name, power state, and IP address of each VM.

    .EXAMPLE
        Get-VmwareVMStatus

        This example retrieves the status of all VMware virtual machines and displays
        the name, power state, and IP address of each VM.

    .INPUTS
        None

    .OUTPUTS
        System.Management.Automation.PSCustomObject
    #>
    $vms = vmrun list | Where-Object { $_ -match 'vmx' } | ForEach-Object {
        $vmxPath = $_ -replace '^.*\s', ''
        $vmxPath
    }
    foreach ($vm in $vms) {
        $vmIp = vmrun getGuestIPAddress $vm
        $vmName = vmrun -gu vagrant -gp vagrant readVariable $vm guestEnv computername
        [PSCustomObject]@{
            Name      = $vmName
            IPAddress = $vmIp
            VmxFile   = $vm
        }
    }
}

Register-ArgumentCompleter -Native -CommandName az -ScriptBlock {
    param($commandName, $wordToComplete, $cursorPosition)
    $completion_file = New-TemporaryFile
    $env:ARGCOMPLETE_USE_TEMPFILES = 1
    $env:_ARGCOMPLETE_STDOUT_FILENAME = $completion_file
    $env:COMP_LINE = $wordToComplete
    $env:COMP_POINT = $cursorPosition
    $env:_ARGCOMPLETE = 1
    $env:_ARGCOMPLETE_SUPPRESS_SPACE = 0
    $env:_ARGCOMPLETE_IFS = "`n"
    $env:_ARGCOMPLETE_SHELL = 'powershell'
    az 2>&1 | Out-Null
    Get-Content $completion_file | Sort-Object | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, "ParameterValue", $_)
    }
    Remove-Item $completion_file, Env:\_ARGCOMPLETE_STDOUT_FILENAME, Env:\ARGCOMPLETE_USE_TEMPFILES, Env:\COMP_LINE, Env:\COMP_POINT, Env:\_ARGCOMPLETE, Env:\_ARGCOMPLETE_SUPPRESS_SPACE, Env:\_ARGCOMPLETE_IFS, Env:\_ARGCOMPLETE_SHELL
}

function Convert-ToLocalTime {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [datetime]$UTCTime
    )

    $localTime = $UTCTime.ToLocalTime()
    return $localTime
}

function Convert-ToUTC {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [datetime]$LocalTime = (Get-Date)
    )

    $utcTime = $LocalTime.ToUniversalTime()
    return $utcTime
}

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

function Convert-CharToInt {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [char[]] $Char
    )
    $Output = @()
    foreach ($c in $Char) {
        $Output += [int]$c
    }
    return $Output
}


function Get-WinEventIDDescription {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [int[]] $EventId,

        [Parameter(Mandatory = $false)]
        [string] $ProviderName,

        [Parameter(Mandatory = $false)]
        [string] $LiteralProviderName,

        [string] $Delimiter
    )

    process {
        if ((-not $ProviderName) -and (-not $LiteralProviderName)) {
            $ProviderName = '*'
            Write-Warning 'No provider name specified. Using wildcard. This may take a while.' -WarningAction Continue
        }
        elseif ($ProviderName) {
            $ProviderName = '*{0}*' -f $ProviderName
        }
        elseif ($LiteralProviderName) {
            $ProviderName = $LiteralProviderName
        }
        $Events = (Get-WinEvent -ListProvider $ProviderName).Events | Where-Object Id -In $EventId
        foreach ($Event in $Events) {
            $Description = (($Event | Select-Object -ExpandProperty Description) -split "`r`n")[0].Trim()
            if ([string]::IsNullOrEmpty($Description)) {
                return [PSCustomObject]@{
                    EventID     = $Event.Id
                    Description = $Description
                }
            }
            else {
                return ('{0}{1}{2}' -f $Event.Id, $Delimiter, $Description)
            }
        }
    }
}

function Set-StarshipConfig {
    $configPath = "~/.config"
    $starshipConfigPath = Join-Path $configPath "starship.toml"

    # Get all .toml files except starship.toml
    $tomlFiles = Get-ChildItem -Path $configPath -Filter "*.toml" | Where-Object { $_.Name -ne "starship.toml" }

    if ($tomlFiles.Count -eq 0) {
        Write-Host "No .toml files found in $configPath."
        return
    }

    # Display menu
    Write-Host "Select a configuration to apply to Starship:"
    $index = 1
    foreach ($file in $tomlFiles) {
        Write-Host "${index}: $($file.Name)"
        $index++
    }

    # Get user choice
    $choice = Read-Host "Enter your choice (1-$($tomlFiles.Count))"
    if (-not [int]::TryParse($choice, [ref]$null) -or $choice -lt 1 -or $choice -gt $tomlFiles.Count) {
        Write-Host "Invalid choice. Operation cancelled."
        return
    }

    # Apply selected configuration
    $selectedFile = $tomlFiles[$choice - 1].FullName
    Get-Content $selectedFile | Out-File $starshipConfigPath -Force
    Write-Host "Configuration applied from $selectedFile to Starship."
}
Set-Alias -Name starconf -Value Set-StarshipConfig


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
function Export-DtmToGot {
    <#
    .SYNOPSIS
        Exports users from Active Directory based on identity or filter criteria
        to a CSV file.

    .DESCRIPTION
        The Export-DtmToGot function retrieves user information from Active
        Directory based on specific identity (e.g., SamAccountName, UserPrincipalName,
        etc.) or a filter string. It then exports the selected properties of the
        retrieved users to a specified CSV file.

    .PARAMETER Identity
        Specifies the identity of the users. Accepts multiple formats like
        SamAccountName, UserPrincipalName, DistinguishedName. You can pass
        an array of identities.

    .PARAMETER Filter
        Specifies a filter in the provider's format or language. The value
        of this parameter qualifies the Path parameter.

    .PARAMETER OutputFile
        Specifies the output file path for the CSV file. If not specified,
        defaults to 'Downloads\dtm.csv' in the user's profile directory.

    .EXAMPLE
        Export-DtmToGot -Identity 'jdoe' -OutputFile 'C:\Users\jdoe\exported_users.csv'

        Exports the user with SamAccountName 'jdoe' to the specified CSV file.

    .EXAMPLE
        Export-DtmToGot -Filter "Department -eq 'Sales'" -OutputFile 'C:\Users\jdoe\sales_users.csv'

        Exports users in the 'Sales' department to the specified CSV file using a filter.

    .INPUTS
        System.String[] (Identity)
        System.String (Filter)

    .OUTPUTS
        CSV file with user details

    .NOTES
        Ensure you have the required permissions to query the Active Directory and write to the output file location.
    #>
    [CmdletBinding(DefaultParameterSetName = 'filter')]
    param(
        [Parameter( Mandatory = $true,
            ParameterSetName = 'identity',
            ValueFromPipeline = $true )]
        [Alias('SamAccountName', 'SamAcct', 'UPN', 'UserPrincipalName', 'DN', 'DistinguishedName')]
        [ValidateNotNullOrEmpty()]
        [string[]] $Identity,

        [Parameter( Mandatory = $true,
            ParameterSetName = 'filter' )]
        [string] $Filter,

        [Parameter( Mandatory = $false )]
        [System.IO.FileInfo] $OutputFile = (Join-Path -Path $env:USERPROFILE -ChildPath 'Downloads\dtm.csv')
    )

    process {
        $Output = [System.Collections.Generic.List[object]]::new()
        $Properties = @(
            'DisplayName'
            'SamAccountName'
            'UserPrincipalName'
            'GivenName'
            'Initials'
            'Surname'
            'EmailAddress'
            'Description'
            'Office'
            'EmployeeId'
            'Title'
            'Department'
            'Manager'
        )
        try {
            if ($PSCmdlet.ParameterSetName -eq 'filter') {
                $Users = Get-ADUser -Filter $Filter -Properties $Properties -ErrorAction Stop | Select-Object $Properties
            }
            else {
                $Users = @()
                foreach ($Id in $Identity) {
                    $Users += Get-ADUser -Identity $Id -Properties $Properties -ErrorAction Stop | Select-Object $Properties
                }
            }

            foreach ($User in $Users) {
                $User.Manager = Get-ADUser -Identity $User.Manager | Select-Object -ExpandProperty SamAccountName
                $User | Add-Member -MemberType NoteProperty -Name RequestorId -Value $env:USERNAME
                $Output.Add($User)
            }

            Write-Information -MessageData "Exporting $($Output.Count) users to $OutputFile" -InformationAction Continue
            $Output | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop -Append
        }
        catch {
            Write-Error $_
        }
    }
}

function Export-PSResource {
    [CmdletBinding()]
    param (
        [Parameter( Mandatory = $true,
            ValueFromPipeline = $true )]
        [String[]] $ResourceName,
        [System.IO.FileInfo] $OutputPath = (Join-Path -Path $env:USERPROFILE -ChildPath 'Downloads'),
        [switch] $Force
    )

    begin {
        if (-not (Test-Path $OutputPath)) {
            if (-not ($Force)) {
                Write-Error ('{0} does not exist. Use -Force to create it.' -f $OutputPath)
                return
            }
            else {
                New-Item -Path $OutputPath -ItemType Directory | Out-Null
            }
        }
    }

    process {
        foreach ($Name in $ResourceName) {
            $Resource = Find-PSResource -Name $Name
            if ([string]::IsNullOrEmpty($Resource)) {
                Write-Error "Resource $Name not found"
                return
            }

            if ([string]::IsNullOrEmpty($Resource.RepositorySourceLocation)) {
                $Source = 'https://www.powershellgallery.com/api/v2'
                Write-Warning -Message "No repository source location found for $Name. Using PSGallery."
            }
            else {
                $Source = $Resource.RepositorySourceLocation
            }

            $URI = '{0}/package/{1}/{2}' -f $Source, $Resource.Name, $Resource.Version
            $OutFile = (Join-Path $OutputPath -ChildPath ('{0}-{1}.nupkg' -f $Resource.Name, $Resource.Version))
            Write-Host "Downloading $URI to $OutFile"
            Invoke-WebRequest -Uri $URI -OutFile $OutFile -ErrorAction SilentlyContinue
        }
    }
}

function Initialize-NewGHRepo {
    [CmdletBinding(DefaultParameterSetName = 'NoDestinationPath')]
    param(
        [string] $Remote,

        [Parameter(ParameterSetName = 'DestinationPath', Mandatory = $true)]
        [Parameter(ParameterSetName = 'NoDestinationPath', Mandatory = $false)]
        [System.IO.DirectoryInfo] $DestinationPath,

        [Parameter(ParameterSetName = 'NoDestinationPath', Mandatory = $true)]
        [Parameter(ParameterSetName = 'DestinationPath', Mandatory = $false)]
        [string] $RepoName
    )

    process {
        # Exit if git isn't found
        try {
            $null = Get-Command -Name git -ErrorAction Stop
        }
        catch {
            Write-Error 'Git is not installed, or not in the PATH.'
            return
        }

        if ($PSCmdlet.ParameterSetName -eq 'NoDestinationPath') {
            $DestinationPath = Get-Item (Join-Path (Get-Location).Path $RepoName)
        }

        if ($PSCmdlet.ParameterSetName -eq 'DestinationPath') {
            if ([string]::IsNullOrEmpty($RepoName)) {
                $RepoName = $DestinationPath | Split-Path -Leaf
            }
        }

        # Create the destination path if it doesn't exist
        if (-not $DestinationPath.Exists) {
            Write-Verbose ('Creating directory {0}' -f $DestinationPath.FullName)
            New-Item -Type Directory -Path $DestinationPath.FullName | Out-Null
        }

        # Change directory to the destination path
        Set-Location -Path $DestinationPath.FullName

        # Exit if the destination path is already a git repository
        if (Get-ChildItem -Path $DestinationPath -Filter '.git' -Hidden -ErrorAction SilentlyContinue) {
            Write-Error ('{0} is already a Git repository.' -f $DestinationPath)
            return
        }

        if ([string]::IsNullOrEmpty($Remote)) {
            $Remote = 'https://github.com/jpbruckler/{0}.git' -f $RepoName
        }

        # Create the README.md file if the directory is empty
        if ((Get-ChildItem | Measure-Object).Count -eq 0) {
            ('# {0}' -f $RepoName) | Out-File -FilePath (Join-Path -Path $DestinationPath -ChildPath 'README.md') -Encoding ascii
        }
    }
    end {
        git init
        git add *
        git commit -m "first commit"
        git branch -M main
        git remote add origin $Remote
        git push -u origin main
        git checkout -b develop
        git push --set-upstream origin develop
    }
}

function Get-PSGalleryModulePackage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$ModuleName,
        [Parameter(Mandatory = $true)]
        [string]$OutputFolder
    )
    process {
        if (-not (Test-Path $OutputFolder)) {
            New-Item -Path $OutputFolder -ItemType Directory | Out-Null
        }
        foreach ($module in $ModuleName) {
            $latestModule = Find-Module $module -Repository PSGallery -ErrorAction Stop
            if ($latestModule) {
                $Uri = ('https://www.powershellgallery.com/api/v2/package/{0}/{1}' -f $module, $latestModule.Version)
                $Path = Join-Path -Path $OutputFolder -ChildPath ('{0}.nupkg' -f $module)
                Write-Verbose "Downloading $module version $($latestModule.Version) from PSGallery"
                Invoke-WebRequest -Uri $Uri -OutFile $Path
            }
            else {
                Write-Warning "Module $module not found on PSGallery"
            }
        }
    }
}

function Update-PSProfile {
    <#
    .SYNOPSIS
        Updates the PowerShell profile with the latest version from a GitHub Gist.

    .DESCRIPTION
        The Update-PowerShellProfile function checks for updates to your PowerShell
        profile stored in a GitHub Gist, and downloads it if there's a newer version
        available. The function assumes you have a file named "Microsoft.PowerShell_profile.ps1"
        in your gist.

    .PARAMETER GistID
        The URL of the raw GitHub Gist containing your PowerShell profile. By
        default this looks for an environment variable of env:PSProfileGistID.

    .EXAMPLE
        Update-PowerShellProfile

        This example checks for updates to the PowerShell profile and downloads
        it if there's a newer version available.

    .INPUTS
        None

    .OUTPUTS
        None

    .NOTES
        If the PowerShell profile is updated, you will need to restart your
        PowerShell session to apply the changes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $GistID = $env:PSProfileGistID
    )

    begin {
        if (-not (Test-Path $PROFILE)) {
            New-Item -Path $PROFILE -ItemType File
        }
    }

    process {
        $ProfileHash = (Get-FileHash $PROFILE).Hash
        $GitUrl = ('https://gist.github.com/{0}.git' -f $GistID)
        $profilePath = $PROFILE
        $tempFile = Join-Path -Path $env:TEMP -ChildPath $GistID

        try {
            # Delete the temporary folder if it exists
            if (Test-Path $tempFile) {
                Remove-Item -Path $tempFile -Recurse -Force
            }

            # Checkout the latest version of the profile from GitHub
            $std0 = (Join-Path -Path $env:temp -ChildPath 'stdout.txt')
            $std1 = (Join-Path -Path $env:temp -ChildPath 'stderr.txt')
            Start-Process git.exe -ArgumentList "clone $GitUrl $tempFile" -Wait -NoNewWindow -RedirectStandardOutput $std0 -RedirectStandardError $std1 -ErrorAction Stop

            # Compare hashes
            if ($ProfileHash -ne (Get-FileHash $tempFile\Microsoft.PowerShell_profile.ps1).Hash) {
                # Copy the updated profile to the PowerShell profile path
                Copy-Item -Path $tempFile\Microsoft.PowerShell_profile.ps1 -Destination $profilePath -Force

                Write-Host "PowerShell profile updated. Restart your PowerShell session to apply the changes."
            }
            else {
                Write-Host "No updates found."
            }
        }
        catch {
            Write-Warning "Unable to delete temporary folder $tempFile"
        }
        finally {
            Remove-Item $std0 -Force
            Remove-Item $std1 -Force
            Remove-Item $tempFile -Recurse -Force
        }
    }
}

function Get-DirectorySize {
    <#
    .SYNOPSIS
        Get the total size of files within a directory.

    .DESCRIPTION
        The Get-DirectorySize function calculates the total size of all files
        within a specified directory and its subdirectories. The result can be
        displayed in bytes, kilobytes, megabytes, or gigabytes.

    .PARAMETER Path
        The path to the directory you want to calculate the size for. Defaults
        to the current directory if not specified.

    .PARAMETER InType
        The unit of measurement for the directory size. Supported values are B
        (bytes), KB (kilobytes), MB (megabytes), and GB (gigabytes). Default is MB.

    .EXAMPLE
        Get-DirectorySize -Path "C:\Temp" -In GB

        This example calculates the total size of all files within the C:\Temp
        directory and its subdirectories, and displays the result in gigabytes.

    .INPUTS
        System.String

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 0)]
        [string] $Path = '.',

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet('B', 'KB', 'MB', 'GB')]
        [string] $In = "MB"
    )

    process {
        $colItems = (Get-ChildItem -Path $Path -Recurse -File | Measure-Object -Property Length -Sum)

        $sizeInBytes = $colItems.Sum
        $formattedSize = switch ($In) {
            "GB" { "{0:N2} GB" -f ($sizeInBytes / 1GB) }
            "MB" { "{0:N2} MB" -f ($sizeInBytes / 1MB) }
            "KB" { "{0:N2} KB" -f ($sizeInBytes / 1KB) }
            "B" { "{0:N2} B" -f $sizeInBytes }
        }

        return $formattedSize
    }
}


function ConvertTo-JsonPretty {
    <#
    .SYNOPSIS
        Convert an object to a pretty-printed JSON string.

    .DESCRIPTION
        The ConvertTo-JsonPretty function takes an input object, converts it to JSON, and then formats the JSON string with proper indentation and line breaks for easier readability.

    .PARAMETER InputObject
        The object to be converted to a pretty-printed JSON string.

    .EXAMPLE
        $data = @{
            Name = "John Doe"
            Age = 30
            City = "New York"
        }

        ConvertTo-JsonPretty -InputObject $data

        This example converts the $data object to a pretty-printed JSON string.

    .INPUTS
        System.Management.Automation.PSObject

    .OUTPUTS
        System.String
    #>
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]
        [PSObject] $InputObject
    )

    process {
        $InputObject | ConvertTo-Json -Depth 100 | Set-Content -Path 'temp.json'
        Get-Content -Path 'temp.json' -Raw | ConvertFrom-Json | ConvertTo-Json -Depth 100
        Remove-Item -Path 'temp.json' -Force
    }
}







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





function New-RandomPassword {
    <#
    .SYNOPSIS
        Generates a random password with specified requirements.

    .DESCRIPTION
        The New-RandomPassword function generates a random password with a given length and character set.
        It ensures that the password contains at least one uppercase character, one lowercase character,
        one special character, and one number. It also allows excluding specific characters from the password.

    .PARAMETER Length
        Specifies the length of the generated password. The default length is 24 characters.

    .PARAMETER ExcludedCharacters
        An array of characters that should not be included in the generated password.
        The default excluded characters are: "`", "|", and "\".

    .EXAMPLE
        New-RandomPassword -Length 16

        Generates a random password with 16 characters, containing at least one uppercase character,
        one lowercase character, one special character, and one number.

    .EXAMPLE
        New-RandomPassword -ExcludedCharacters @("A", "1", "@")

        Generates a random password that does not contain the characters "A", "1", or "@".

    .NOTES
        The generated password may not be cryptographically secure.
        Use the function with caution when generating hi$DriveName-security passwords.

    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$Length = 24,

        [Parameter(Mandatory = $false)]
        [string[]]$ExcludedCharacters = @('`', '|', '\'),

        [switch] $AsSecureString
    )

    process {
        $upperCase = 'ABCDEF$DriveNameIJKLMNOPQRSTUVWXYZ'
        $lowerCase = 'abcdef$DriveNameijklmnopqrstuvwxyz'
        $digits = '0123456789'
        $specialChars = '!@#$%^&*()-_=+[]{};:,.<>?/'

        $password = @()

        $password += Get-Random -InputObject $upperCase.ToCharArray()
        $password += Get-Random -InputObject $lowerCase.ToCharArray()
        $password += Get-Random -InputObject $digits.ToCharArray()
        $password += Get-Random -InputObject $specialChars.ToCharArray()

        $remainingLength = $Length - 4
        $allChars = $upperCase + $lowerCase + $digits + $specialChars

        # Remove excluded characters from the character set
        $allowedChars = $allChars.ToCharArray() | Where-Object { $_ -notin $ExcludedCharacters }

        for ($i = 0; $i -lt $remainingLength; $i++) {
            $password += Get-Random -InputObject $allowedChars
        }

        # Shuffle the password characters
        $password = $password | Get-Random -Count $password.Count

        # Convert the password array to a string
        if ($AsSecureString) {
            $output = ConvertTo-SecureString (-join $password) -AsPlainText -Force
        }
        else {
            $output = -join $password
        }

        Write-Output $output
    }
}

function New-DcrXPathFilter {

    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [int[]] $EventIds,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $LogName,

        [ValidateSet('or', 'and')]
        [string] $Operator = 'or',

        [switch] $LookupLogName
    )

    if ($LookupLogName) {
        $LogName = (Get-WinEvent -ListLog * -ea SilentlyContinue | Where-Object LogName -Like "*$LogName*").LogName
        if (-not $LogName) {
            throw 'Log name not found. Try running the following: (Get-WinEvent -ListLog * | Where-Object LogName -like "*$LogName*").LogName'
        }
    }





    if ($EventIds.Count -eq 1) {
        return "$LogName!*[System[EventID=${EventIds}]]"
    }
    elseif ($EventIds.Count -gt 20) {
        $output = @()
        $chunks = Split-Array -InputArray $EventIds -MaximumElements 20
        foreach ($chunk in $chunks) {
            $output += New-DcrXPathFilter -EventIds $chunk -LogName $LogName -Operator $Operator
        }
        #$output += New-DcrXPathFilter -EventIds $EventIds[0..19] -LogName $LogName -Operator $Operator
        #$output += New-DcrXPathFilter -EventIds $EventIds[20..($EventIds.Count - 1)] -LogName $LogName -Operator $Operator
        return $output
    }
    else {
        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append(('{0}!*[System[' -f $LogName))
        $limit = $EventIds.Count - 1
        for ($i = 0; $i -lt $limit; $i++) {
            $eid = $EventIds[$i]
            $null = $sb.Append("(EventID=$eid) $Operator ")
        }
        # Append the last EventID without the 'or'
        $eid = $EventIds[-1]
        $null = $sb.Append("(EventID=$eid)")
        $null = $sb.Append(']]')
        return $sb.ToString()
    }
}

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

function Get-AttackSurfaceReductionConfig {
    $guidLookup = @{
        "56a863a9-875e-4185-98a7-b882c64b5ce5" = "Block abuse of exploited vulnerable signed drivers"
        "7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c" = "Block Adobe Reader from creating child processes"
        "d4f940ab-401b-4efc-aadc-ad5f3c50688a" = "Block all Office applications from creating child processes"
        "9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2" = "Block credential stealing from the Windows local security authority subsystem (lsass.exe)"
        "be9ba2d9-53ea-4cdc-84e5-9b1eeee46550" = "Block executable content from email client and webmail"
        "01443614-cd74-433a-b99e-2ecdc07bfc25" = "Block executable files from running unless they meet a prevalence, age, or trusted list criterion"
        "5beb7efe-fd9a-4556-801d-275e5ffc04cc" = "Block execution of potentially obfuscated scripts"
        "d3e037e1-3eb8-44c8-a917-57927947596d" = "Block JavaScript or VBScript from launching downloaded executable content"
        "3b576869-a4ec-4529-8536-b80a7769e899" = "Block Office applications from creating executable content"
        "75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84" = "Block Office applications from injecting code into other processes"
        "26190899-1602-49e8-8b27-eb1d0a1ce869" = "Block Office communication application from creating child processes"
        "e6db77e5-3df2-4cf1-b95a-636979351e5b" = "Block persistence throu$DriveName WMI event subscription"
        "d1e49aac-8f56-4280-b9ba-993a6d77406c" = "Block process creations originating from PSExec and WMI commands"
        "b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4" = "Block untrusted and unsigned processes that run from USB"
        "92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b" = "Block Win32 API calls from Office macros"
        "c1db55ab-c21a-4637-bb3f-a12568109d35" = "Use advanced protection against ransomware"
    }

    $actionLookup = @{
        0 = "Off"
        1 = "Block"
        2 = "Audit"
        5 = "Not Configured"
        6 = "Warn"
    }

    Get-MpPreference | ForEach-Object {
        for ($i = 0; $i -lt $_.AttackSurfaceReductionRules_Ids.Count; $i++) {
            $id = $_.AttackSurfaceReductionRules_Ids[$i]
            [PSCustomObject]@{
                Id       = $id
                Name     = $guidLookup[$id]
                ActionId = $_.AttackSurfaceReductionRules_Actions[$i]
                Action   = $actionLookup[[int]$_.AttackSurfaceReductionRules_Actions[$i]]
            }
        }
    }
}

function Get-PercentOf {
    param(
        [int] $Total,
        [int] $Part
    )

    $percent = ($Part / $Total) * 100
    return $percent
}


function Test-LDAPPorts {
    [CmdletBinding()]
    param(
        [string] $ServerName,
        [int] $Port
    )
    if ($ServerName -and $Port -ne 0) {
        try {
            $LDAP = "LDAP://" + $ServerName + ':' + $Port
            $Connection = [ADSI]($LDAP)
            $Connection.Close()
            return $true
        }
        catch {
            if ($_.Exception.ToString() -match "The server is not operational") {
                Write-Warning "Can't open $ServerName`:$Port."
            }
            elseif ($_.Exception.ToString() -match "The user name or password is incorrect") {
                Write-Warning "Current user ($Env:USERNAME) doesn't seem to have access to to LDAP on port $Server`:$Port"
            }
            else {
                Write-Warning -Message $_
            }
        }
        return $False
    }
}
Function Test-LDAP {
    [CmdletBinding()]
    param (
        [alias('Server', 'IpAddress')][Parameter(Mandatory = $True)][string[]]$ComputerName,
        [int] $GCPortLDAP = 3268,
        [int] $GCPortLDAPSSL = 3269,
        [int] $PortLDAP = 389,
        [int] $PortLDAPS = 636
    )
    # Checks for ServerName - Makes sure to convert IPAddress to DNS
    foreach ($Computer in $ComputerName) {
        [Array] $ADServerFQDN = (Resolve-DnsName -Name $Computer -ErrorAction SilentlyContinue)
        if ($ADServerFQDN) {
            if ($ADServerFQDN.NameHost) {
                $ServerName = $ADServerFQDN[0].NameHost
            }
            else {
                [Array] $ADServerFQDN = (Resolve-DnsName -Name $Computer -ErrorAction SilentlyContinue)
                $FilterName = $ADServerFQDN | Where-Object { $_.QueryType -eq 'A' }
                $ServerName = $FilterName[0].Name
            }
        }
        else {
            $ServerName = ''
        }

        $GlobalCatalogSSL = Test-LDAPPorts -ServerName $ServerName -Port $GCPortLDAPSSL
        $GlobalCatalogNonSSL = Test-LDAPPorts -ServerName $ServerName -Port $GCPortLDAP
        $ConnectionLDAPS = Test-LDAPPorts -ServerName $ServerName -Port $PortLDAPS
        $ConnectionLDAP = Test-LDAPPorts -ServerName $ServerName -Port $PortLDAP

        $PortsThatWork = @(
            if ($GlobalCatalogNonSSL) { $GCPortLDAP }
            if ($GlobalCatalogSSL) { $GCPortLDAPSSL }
            if ($ConnectionLDAP) { $PortLDAP }
            if ($ConnectionLDAPS) { $PortLDAPS }
        ) | Sort-Object
        [pscustomobject]@{
            Computer           = $Computer
            ComputerFQDN       = $ServerName
            GlobalCatalogLDAP  = $GlobalCatalogNonSSL
            GlobalCatalogLDAPS = $GlobalCatalogSSL
            LDAP               = $ConnectionLDAP
            LDAPS              = $ConnectionLDAPS
            AvailablePorts     = $PortsThatWork -join ','
        }
    }
}

function ConvertTo-SafeUrl {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$Urls
    )

    process {
        foreach ($Url in $Urls) {
            # Replace 'http' with 'hxxp'
            $safeUrl = $Url -replace 'http', 'hxxp'
            # Replace '.' with '[.]', except for the first occurrence after 'hxxp://'
            $safeUrl = $safeUrl -replace '://', '://' -replace '\.', '[.]'
            Write-Output $safeUrl
        }
    }
}

function ConvertFrom-SafeIoc {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Iocs
    )

    process {
        foreach ($Ioc in $Iocs) {
            # Replace 'hxxp' with 'http'
            $url = $Ioc -replace 'hxxp', 'http'
            # Replace '[.]' with '.'
            $url = $url -replace '\[\.\]', '.'
            Write-Output $url
        }
    }
}


function Format-WinLogonEvents {
    [CmdletBinding(DefaultParameterSetName = 'Search')]
    param(
        [Parameter( Mandatory = $true,
            ValueFromPipeline = $true,
            ParameterSetName = 'Events')]
        [System.Diagnostics.Eventing.Reader.EventLogRecord[]]$Events,
        [Parameter( Mandatory = $false,
            ParameterSetName = 'Search')]
        [string]$ComputerName,
        [Parameter( Mandatory = $false,
            ParameterSetName = 'Search')]
        [string]$InputFile,
        [Parameter( Mandatory = $false,
            ParameterSetName = 'Search')]
        [datetime]$StartTime,
        [Parameter( Mandatory = $false,
            ParameterSetName = 'Search')]
        [datetime]$EndTime,
        [Parameter( Mandatory = $false,
            ParameterSetName = 'Search')]
        [int]$MaxEvents
    )

    process {
        $OutputEvents = [System.Collections.Generic.List[pscustomobject]]::new()
        $FilterHashtable = @{
            LogName = 'Security'
            ID      = 4624, 4634, 4647
        }
        if ($StartTime) { $FilterHashtable.StartTime = $StartTime }
        if ($EndTime) { $FilterHashtable.EndTime = $EndTime }
        if ($MaxEvents) { $FilterHashtable.MaxEvents = $MaxEvents }
        if ($InputFile) {
            $FilterHashtable.Path = $InputFile
            $FilterHashtable.Remove('LogName')
        }
        if ($PSCmdlet.ParameterSetName -eq 'Events') {
            $Events = $Events | Where-Object { $_.Id -eq 4624 -or $_.Id -eq 4634 -or $_.Id -eq 4647 }
        }
        else {
            $GetWinEventSplat = @{
                FilterHashTable = $FilterHashtable
            }

            if ($null -ne $ComputerName) {
                $GetWinEventSplat.ComputerName = $ComputerName
            }
            if ($null -ne $MaxEvents) {
                $GetWinEventSplat.MaxEvents = $MaxEvents
            }
            $Events = Get-WinEvent @GetWinEventSplat
        }

        $Events | ForEach-Object {
            $Eventx = $_
            $EventXml = [xml]$Event.ToXml()

            if ($EventId = 4624) {
                $TargetUserName = $EventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' } | Select-Object -ExpandProperty '#text'
            }

            $OutEvent = [PSCustomObject] @{
                TimeCreated = $Eventx.TimeCreated
                EventID     = $Eventx.Id
            }
        }
    }

}

function Rename-PSIAMUser {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Hi$DriveName')]
    param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true)]
        [string] $UserName,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true)]
        [alias('NewFirstName')]
        [string] $NewGivenName,
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true)]
        [alias('NewLastName')]
        [string] $NewSurname,
        [pscredential] $Credential
    )

    begin {
        $ti = (Get-Culture).TextInfo
    }

    process {
        try {
            $adUser = Get-ADUser -Identity $UserName -Properties EmailAddress -Credential $Credential -ErrorAction Stop
        }
        catch {
            Write-Error "User $UserName not found."
            return
        }

        $GivenName = $ti.ToTitleCase($NewGivenName)
        $Surname = if (-not ([string]::IsNullOrWhiteSpace($NewSurname))) {
            Write-Host 'here'
            $ti.ToTitleCase($NewSurname)
        }
        else {
            $adUser.Surname
        }
        $DisplayName = '{0} {1}' -f $GivenName, $Surname
        $newEmailAddress = $adUser.UserPrincipalName -replace $adUser.GivenName, $GivenName
        $proxyAddresses = 'smtp:{0}, SMTP:{1}' -f $newEmailAddress, $adUser.EmailAddress

        $setAduserParams = @{}
        $setAduserParams.GivenName = $GivenName
        $setAduserParams.Surname = $Surname
        $setAduserParams.Add = @{ProxyAddresses = $proxyAddresses }
        $setAduserParams.DisplayName = $DisplayName

        if ($PSCmdlet.ShouldProcess($UserName, "Set name to $DisplayName and add email address $newEmailAddress")) {
            $adUser | Set-ADUser @setAduserParams -Credential $Credential
            Get-ADUser -Identity $UserName -Credential $Credential -Properties EmailAddress, ProxyAddresses, DisplayName |
            Select-Object Name, GivenName, Surname, DisplayName, EmailAddress, ProxyAddresses
        }
    }
}

function ConvertPFX-ToPem {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ParameterSetName = 'Path')]
        [String] $Path,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ParameterSetName = 'LiteralPath')]
        [System.IO.FileInfo] $LiteralPath,

        [Parameter(Mandatory = $true)]
        [securestring] $Password,

        [Parameter(Mandatory = $false)]
        [System.IO.DirectoryInfo] $Output = (Get-Location).Path,

        [switch] $Dearmor
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $ImportPath = [System.IO.FileInfo] (Resolve-Path -Path $Path).Path
        }
        else {
            $ImportPath = $LiteralPath
        }


        $keyFilePath = Join-Path $Output.FullName $ImportPath.BaseName + '_key.pem'
        $certFilePath = Join-Path $Output.FullName $ImportPath.BaseName + '_cert.pem'
        $certChainFilePath = Join-Path $Output.FullName $ImportPath.BaseName + '_chain.pem'

        $Password = $Password | ConvertFrom-SecureString -AsPlainText
        $pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ImportPath,
                    $Password, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)


        # Export the private key
        $privateKey = $pfx.PrivateKey
        $privateKeyBytes = $privateKey.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs8, $null)
        $privateKeyPem = [System.Convert]::ToBase64String($privateKeyBytes)

        # Export the certificate
        $certBytes = $pfx.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        $certPem = [System.Convert]::ToBase64String($certBytes)

        # Export the full chain if there are intermediate certificates
        $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
        $chain.Build($pfx)
        $certChainPem = $chain.ChainElements | ForEach-Object {
            $cert = $_.Certificate
            $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
            [System.Convert]::ToBase64String($certBytes)
        }

        # Write the full chain to PEM file if applicable
        if ($chain.ChainElements.Count -gt 1) {
            $certChainPemContent = $certChainPem | ForEach-Object { "-----BEGIN CERTIFICATE-----`n$($_ -replace '(.{64})', '$1`n')`n-----END CERTIFICATE-----" }
        }

        $privateKeyPemContent = "-----BEGIN PRIVATE KEY-----`n$($privateKeyPem -replace '(.{64})', '$1`n')`n-----END PRIVATE KEY-----"
        $certPemContent = "-----BEGIN CERTIFICATE-----`n$($certPem -replace '(.{64})', '$1`n')`n-----END CERTIFICATE-----"

        if ($Dearmor) {
            # Write the private key in binary format
            [System.IO.File]::WriteAllBytes($keyFilePath, $privateKeyBytes)

            # Write the certificate in binary format
            [System.IO.File]::WriteAllBytes($CertFilePath, $certBytes)

            # Write the full chain in binary format if applicable
            if ($chain.ChainElements.Count -gt 1) {
                $certChainBytes = $certChainPem | ForEach-Object { [System.IO.File]::WriteAllBytes($certChainFilePath, $_) }
            }
        }
        else {
            Set-Content -Path $certFilePath -Value $certPemContent
            Set-Content -Path $keyFilePath -Value $privateKeyPemContent
            Set-Content -Path $certChainFilePath -Value ($certChainPemContent -join "`n")
        }
    }
}

function Get-PrismaGPGateway {
    [CmdletBinding()]
    param (
        [string] $ApiKey = '29cx2plida__Kpl5LwtTji7tWzkVHskw03QejwI7nuow9cmP4rCj'
    )

    $uri = "https://api.prod.datapath.prismaaccess.com/getPrismaAccessIP/v2"


    $headers = @{}
    $headers.Add("Accept", "*/*")
    $headers.Add("User-Agent", "Thunder Client (https://www.thunderclient.com)")
    $headers.Add("header-api-key", $ApiKey)
    $headers.Add("Content-Type", "application/json")

    $reqUrl = 'https://api.prod.datapath.prismaaccess.com/getPrismaAccessIP/v2'
    $body = '{
        "addrType": "all",
        "location": "all",
        "serviceType": "gp_gateway"
    }'

    $response = Invoke-RestMethod -Uri $reqUrl -Method Post -Headers $headers -ContentType 'application/json' -Body $body | Select-Object -ExpandProperty result
    $subnets = $response | Select-Object -ExpandProperty zone_subnet -Unique

    $sorted = $subnets | Sort-Object {
        $ipAddress = $_.split('/')[0]
        $octets = $ipAddress.Split('.') | ForEach-Object { [int]$_ }
        $value = ($octets[0] * [math]::Pow(256, 3)) + ($octets[1] * [math]::Pow(256, 2)) + ($octets[2] * 256) + $octets[3]
        $value
    }
    return $sorted
}
#endregion
