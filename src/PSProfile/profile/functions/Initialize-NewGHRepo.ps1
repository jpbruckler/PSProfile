function Initialize-NewGHRepo {
    <#
    .SYNOPSIS
        Initialize a new GitHub repository.
    .DESCRIPTION
        This function initializes a new GitHub repository at the given path. If
        no path is provided, the new repository will be created in the current
        directory using the value of the parameter RepositoryName as the name of
        the subdirectory.
    .PARAMETER Remote
        The remote URL of the GitHub repository. If not provided, the value of
        the environment variable $env:PSPROFILE_GITHUB_REMOTE will be used. This
        environment variable should be set in your profile script.
    .PARAMETER Path
        The path where the new repository will be created. If not provided, the
        current directory will be used.
    .PARAMETER RepositoryName
        The name of the new repository. This will be used as the name of the
        subdirectory in the path where the repository will be created.
    .EXAMPLE
        Initialize-NewGHRepo -RepositoryName 'MyNewRepo'

        This will create a new GitHub repository named 'MyNewRepo' in the current
        directory using the value of $env:PSPROFILE_GITHUB_REMOTE/MyNewRepo.git
        as the remote URL.
    .EXAMPLE
        Initialize-NewGHRepo -RepositoryName 'MyNewRepo' -Path 'C:\Projects'

        This will create a new GitHub repository named 'MyNewRepo' in the directory
        'C:\Projects' using the value of $env:PSPROFILE_GITHUB_REMOTE/MyNewRepo.git
        as the remote URL.

    .EXAMPLE
        Initialize-NewGHRepo -RepositoryName 'MyNewRepo' -Remote 'https://github.com/jpbruckler/'

        This will create a new GitHub repository named 'MyNewRepo' in the current
        directory with a remote URL of https://github.com/jpbruckler/MyNewRepo.git.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $false)]
        [Alias('Remote', 'Upstream', 'Origin')]
        [string] $Remote = $env:PSPROFILE_GITHUB_REMOTE,

        [Parameter(Mandatory = $false)]
        [System.IO.DirectoryInfo] $Path = (Get-Location).Path,

        [Parameter(Mandatory = $true)]
        [Alias('Name')]
        [string] $RepositoryName
    )

    begin {
        # Exit if git isn't found
        try {
            $null = Get-Command -Name git -ErrorAction Stop
        }
        catch {
            Write-Error 'Git is not installed, or not in the PATH.'
            return
        }

        # Exit if the destination path is already a git repository
        if (Get-ChildItem -Path $Path -Filter '.git' -Hidden -ErrorAction SilentlyContinue) {
            Write-Error ('{0} is already a Git repository, exiting.' -f $Path)
            return
        }
    }

    process {

        $Remote = $Remote.Trim('/')

        if (-not ($Remote.EndsWith('.git'))) {
            if ($Remote.EndsWith($RepositoryName)) {
                Write-Verbose 'Adding .git to end of remote URL.'
                $Remote = '{0}.git' -f $Remote, $RepositoryName
            }
            else {
                Write-Verbose 'Adding {0}.git to end of remote URL.' -f $RepositoryName
                $Remote = '{0}/{1}.git' -f $Remote, $RepositoryName
            }

            Write-Verbose ('Remote URL: {0}' -f $Remote)
        }

        $ShouldProcessMsg = 'Initialize a new GitHub repository in {0} with a remote URL of {1}' -f $Path, $Remote
        if ($PSCmdlet.ShouldProcess($RepositoryName, $ShouldProcessMsg)) {
            Write-Verbose ('Initializing a new GitHub repository in {0}' -f $Path)
            # Create the destination path if it doesn't exist
            if (-not $Path.Exists) {
                Write-Verbose ('Creating directory {0}' -f $Path.FullName)
                $null = New-Item -Type Directory -Path $Path.FullName
            }

            # Create the README.md file if the directory is empty
            if ((Get-ChildItem | Measure-Object).Count -eq 0) {
                ('# {0}' -f $RepositoryName) | Out-File -FilePath (Join-Path -Path $Path -ChildPath 'README.md') -Encoding ascii
            }

            # run the git commands
            Set-Location -Path $Path
            git init
            git add *
            git commit -m "first commit"
            git branch -M main
            git remote add origin $Remote
            git push -u origin main
            git checkout -b develop
            git push --set-upstream origin develop

            Write-Information ('GitHub repository {0} initialized in {1}' -f $RepositoryName, $Path) -InformationAction Continue
            Write-Information ("`tRemote URL: {0}" -f $Remote) -InformationAction Continue
            Write-Information ("`tCurrent branch: develop") -InformationAction Continue
            Write-Information ("`tA default README.md file has been created.") -InformationAction Continue
        }
    }
}