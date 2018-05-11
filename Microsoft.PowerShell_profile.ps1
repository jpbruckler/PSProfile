# Setup PSDrives
if (-not (Test-Path gh:\)) {
    New-PSDrive -Name GH -PSProvider FileSystem -Root (Join-Path $env:USERPROFILE 'Documents\Github')
}

# Variables
$here               = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProfilePath        = Split-Path $Profile -Parent

# Path manipulation
# Add paths to the array to add paths to the.. to the path...
$PathsToAdd = @(
    (Join-Path ${env:ProgramFiles(x86)} 'nmap')             # Nmap
    (Join-Path $env:ProgramFiles 'MongoDB\Server\3.4\bin')  # Mongo
    (Join-Path $env:ProgramFiles 'Beyond Compare 3')        # BeyondCompare
    (Join-Path $env:ProgramFiles 'Sysinternals')            # SysInternals Tools
    (Join-Path $env:ProgramFiles 'Microsoft VS Code')       # VS Code
    (Join-Path $env:LOCALAPPDATA 'Programs\Git\cmd')        # git
)
foreach ($Path in $PathsToAdd) {
    $NotInPath  = $env:Path -split ';' -notcontains $Path
    $TestPath   = Test-Path $Path -ErrorAction SilentlyContinue
    if ($NotInPath -and $TestPath) {
        $env:Path = '{0};{1}' -f $env:Path, $Path
    }
}

# Some of my machines for some reason have a messed up PSModulePath, this will fix it.
$UserModulePath = (Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules')
if (($env:PSModulePath -split ';') -notcontains $UserModulePath) {
    $env:PSModulePath   = '{0};{1}' -f $env:PSModulePath, $UserModulePath
}

# Aliases
Set-Alias -Name scp -Value $((Join-Path ${env:ProgramFiles(x86)} 'Centrify\Centrify PuTTY\pscp.exe' ) )

# Modules
Import-Module posh-git

# Functions
function src { & $Profile }
function which($name) { Get-Command $name | Select-Object Definition }
function touch($file) { "" | Out-File $file -Encoding ASCII }
function Find-MSDN
{
    param([string] $SearchTerm)

    process { 
        $url = 'https://social.msdn.microsoft.com/Search/en-US?query={0}' -f $SearchTerm
        Start-Process $url
    }
}

function Open-MITRE 
{
    <#
    .SYNOPSIS
        Opens MITRE ATT&CK wiki page for the given T code in the default browser.
    .DESCRIPTION
        For a given technique code (Tnnnn), opens the wiki page for the corresponding 
        MITRE ATT&CK technique in the default browser.
    .PARAMETER TechniqueCode
        The ATT&CK technique code. This can either be a 4 digit number, or the fully
        qualified technique ID (e.g. T1120) of a documented MITRE ATT&CK technique.
        
        The leading T can be left off, it will be added by the function if missing.
    .INPUTS
        String
    .OUTPUTS
        Null
    .EXAMPLE
        PS C:\> Find-MITRE -TechniqueCode T1120

        The above example will open a tab in the default browser to https://attack.mitre.org/wiki/Technique/T1120.

    .EXAMPLE
        PS C:\> 1110..1120 | Find-MITRE

        The above example will open up 10 tabs, 1 for each T code. Note that the T is left 
        off the technique code.
    #>
    param(
        [Parameter( Mandatory,
                    ValueFromPipeline )]    
        [Alias('TCode')]
        [string] $TechniqueCode
    )

    process {
        if (-not ($TechniqueCode.StartsWith('T'))) {
            $TechniqueCode = 'T{0}' -f $TechniqueCode
        }

        $url = ('https://attack.mitre.org/wiki/Technique/{0}' -f $TechniqueCode)
        Start-Process $url
    }
}

function Set-HttpProxy
{
    param(
        [Parameter( Mandatory )]
        [string] $ProxyServer,
        [int] $Port = 8080,
        [System.Management.Automation.PSCredential] $Credential
    )
    
    begin {
        Add-Type -AssemblyName System.Web
    }
  
    process { 
        if ($Credential) {
            $UserName = $Credential.GetNetworkCredential().UserName
            $Password = [System.Web.HttpUtility]::UrlEncode($Credential.GetNetworkCredential().Password)
            $ProxyString = 'http://{0}:{1}@{2}:{3}' -f $UserName,$Password,$ProxyServer,$Port

            $UserName, $Password = $null
        }
        else {
            $ProxyString = 'http://{0}:{1}' -f $ProxyServer,$Port
        }
        
        $env:HTTP_PROXY     = $ProxyString
        $env:HTTPS_PROXY    = $ProxyString
        $env:http_proxy     = $ProxyString
        $env:https_proxy    = $ProxyString
    }
}

function Clear-HttpProxy
{
  $env:HTTP_PROXY   = $null
  $env:HTTPS_PROXY  = $null
  $env:http_proxy   = $null
  $env:https_proxy  = $null
}

function ConvertTo-Base64
{
  param( 
    [Parameter( Mandatory,
                ValueFromPipeline )]
    [string[]] $String
  )
  
  process {
    $Bytes  = [System.Text.Encoding]::Unicode.GetBytes($String)
    $Output = [System.Convert]::ToBase64String($Bytes)
    Write-Output $Output
  }
}

function ConvertFrom-Base64
{
  param( 
    [Parameter( Mandatory,
                ValueFromPipeline )]
    [Alias('String','Base64')]
    [string[]] $Base64String
  )
  
  process {
    $Bytes = [System.Convert]::FromBase64String($Base64String)
    $Output = [System.Text.Encoding]::Unicode.GetString($Bytes)
    Write-Output $Output    
  }
}

function Start-TestShell
{
    param( 
        [string] $TargetDir = $PWD
    )

    Push-Location $TargetDir
    $ScriptBlock = {
        function prompt {
            $Text = '[TST] {0}{1}' -f $($PWD.Path.Substring($PWD.Path.LastIndexOf('\'))), $('>' * ($nestedPromptLevel + 1))
            Write-Host -Object $Text -NoNewline -ForegroundColor Green -BackgroundColor DarkGreen
            return ' '
        }
    }
    PowerShell.exe -NoExit -NoProfile -Command $ScriptBlock
}
Set-Alias -Name sts -Value Start-TestShell

function Start-IronKey
{
    [CmdletBinding()]
    param( [switch] $LaunchAll)
    
    process {
        # Get CD-ROM Drive types, because the IronKey unlocker partition
        # is mounted as a CD.
        $Drives = Get-WmiObject Win32_LogicalDisk -Filter DriveType=5

        foreach ($Drive in $Drives)
        {
            $IKeyPath = (Join-Path $Drive.DeviceID 'IronKey.exe')

            Write-Verbose ('Checking drive {0} for IronKey.exe.' -f $Drive.DeviceID)

            if ((Resolve-Path $IKeyPath -ErrorAction SilentlyContinue))
            {
                Write-Verbose ('IronKey.exe found on drive {0}. Launching.' -f $Drive.DeviceID)
                Start-Process $IKeyPath

                if ($LaunchAll -eq $false) { break; }
            }
        }
    }
}
Set-Alias -Name sik -Value Start-IronKey

function Get-TinyUrl
{
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [string] $URL
    )

    $TinyURL = "http://tinyurl.com/api-create.php?url=$URL"
    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadString($TinyURL)
}

function Get-FileHash
{
    param(
        [Parameter(Mandatory=$true,
                    ValueFromPipeline=$true)]
        [string] $FilePath,
        
        [ValidateSet('MD5','RIPEMD160','SHA1','SHA256','SHA384','SHA512')]
        [string] $Algorithm = 'SHA1'
    )

    begin
    {
        $Crypt = [Security.Cryptography.HashAlgorithm]::Create($Algorithm)
    }

    process
    {
        if ([System.IO.File]::Exists($FilePath))
        {
            $Stream = [IO.File]::OpenRead($FilePath)
            [string] $hash = $Crypt.ComputeHash($Stream) | ForEach-Object { '{0:x2}' -f $_ }
            $Stream.Close()
            $Output = [PSCustomObject] @{ 
                        'FilePath' = $FilePath
                        'FileName' = (Split-Path $FilePath -Leaf)
                        'Hash' = ($hash -replace ' ','') 
                        'Algorithm' = $Algorithm
                      }
            $Output
        }
        else
        {
            $Message = "The given file path '$FilePath' is not a valid file or directory."
            $Exception = New-Object InvalidOperationException $Message
            $ErrorID = 'FileIsMissingOrEmpty'
            $ErrorCategory = [Management.Automation.ErrorCategory]::ObjectNotFound
            $Target = $FilePath
            $ErrorRecord = New-Object Management.Automation.ErrorRecord $Exception, $ErrorID, $ErrorCategory, $Target
            Throw $ErrorRecord
        }
    }
}

function Compare-FileHash
{
    param(
        [Parameter(Mandatory=$true,
                    ValueFromPipeline=$true)]
        [string] $FilePath,
        
        [ValidateSet('MD5','RIPEMD160','SHA1','SHA256','SHA384','SHA512')]
        [string] $Algorithm = 'SHA1',
        [string] $CheckSum
    )

    begin
    {
        $Crypt = [Security.Cryptography.HashAlgorithm]::Create($Algorithm)
    }

    process
    {
        if (Test-Path $FilePath)
        {
            $Stream = [IO.File]::OpenRead($FilePath)
            [string] $hash = $Crypt.ComputeHash($Stream) | ForEach-Object { '{0:x2}' -f $_ }
            $Stream.Close()
            $FileHash = ($hash -replace ' ','')
        }
        else
        {
            $Message = "The given file path '$FilePath' is not a valid file or directory."
            $Exception = New-Object InvalidOperationException $Message
            $ErrorID = 'FileIsMissingOrEmpty'
            $ErrorCategory = [Management.Automation.ErrorCategory]::ObjectNotFound
            $Target = $FilePath
            $ErrorRecord = New-Object Management.Automation.ErrorRecord $Exception, $ErrorID, $ErrorCategory, $Target
            Throw $ErrorRecord
        }

        Write-Output ($FileHash -match "^$CheckSum$")
    }
}

Function Compare-GPOSettings
{
    param(
        [Parameter( Mandatory )]
        [string] $ReferenceGPOName,

        [Parameter( Mandatory )]
        [string] $DifferenceGPOName,

        [switch] $Computer
    )

    begin {
        $ModuleExists = Get-Module -Name GroupPolicy
        if ($null -eq $ModuleExists) {
            throw 'GroupPolicy module cannot be found. Install Remote Server Administration Tools for your Operating System version.'
        }
        else {
            Import-Module GroupPolicy
        }
    }

    process {
        try {
            $RefGPO = Get-GPO -Name $ReferenceGPOName
            $DifGPO = Get-GPO -Name $DifferenceGPOName

            $RefXMLFile = ('{0}\{1}.xml' -f $env:TEMP,$ReferenceGPOName)
            $DifXMLFile = ('{0}\{1}.xml' -f $env:TEMP,$DifferenceGPOName)

            $null = $RefGPO.GenerateReportToFile('xml',$RefXMLFile)
            $null = $DifGPO.GenerateReportToFile('xml',$DifXMLFile)

            [xml] $RefGPOXml = Get-Content $RefXMLFile
            [xml] $DifGPOXml = Get-Content $DifXMLFile

            switch ($Computer) {
                $true {
                    $ComputerCompareObjParams = @{
                        ReferenceObject  = ($RefGPOXml.gpo.Computer.ExtensionData.extension.ChildNodes | Select-Object Name,State)
                        DifferenceObject = ($DifGPOXml.gpo.Computer.ExtensionData.extension.ChildNodes | Select-Object Name,State)
                        IncludeEqual     = $true
                        Property         = 'Name'
                    }
                    Compare-Object @ComputerCompareObjParams
                }
                
                $false {
                    $UserCompareObjParams = @{
                        ReferenceObject  = ($RefGPOXml.gpo.User.ExtensionData.extension.ChildNodes | Select-Object Name,State)
                        DifferenceObject = ($DifGPOXml.gpo.User.ExtensionData.extension.ChildNodes | Select-Object Name,State)
                        IncludeEqual     = $true
                        Property         = 'Name'
                    }
                    Compare-Object @UserCompareObjParams
                }
            }
        }
        catch {
            Write-Error $PSItem
        }

    }
}
