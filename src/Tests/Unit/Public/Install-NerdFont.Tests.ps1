# Install-NerdFont.Tests.ps1


# Describe block for grouping related tests
Describe 'Install-NerdFont' {

    # BeforeAll block for setting up common mocks
    BeforeAll {
        $verb = 'Install'
        $fnDir = (Resolve-Path ([System.IO.Path]::Combine('..', '..', 'PSProfile','src','PSProfile','Functions'))).Path
        . (Resolve-Path (Join-Path -Path $fnDir -ChildPath "$verb\Install-NerdFont.ps1")).Path

        # Mock external dependencies
        Mock -CommandName Invoke-RestMethod -MockWith {
            @{
                assets = @(
                    @{name = 'CascadiaCode'; browser_download_url = 'https://example.com/CascadiaCode.zip'},
                    @{name = 'FiraCode'; browser_download_url = 'https://example.com/FiraCode.zip'}
                )
            }
        }

        Mock -CommandName Invoke-WebRequest
        Mock -CommandName Expand-Archive
        Mock -CommandName Install-Font
    }

    Context "Installing one font" {
        BeforeEach {
            Install-NerdFont -NerdFontName 'CascadiaCode'
        }
        # Test for downloading and installing a specific font
        It 'Downloads and installs a specific Nerd Font' {
            Should -Invoke -CommandName Install-Font -Exactly -Times 1
        }
    }

    Context "Installing all fonts" {
        BeforeEach {
            Install-NerdFont -All
        }
        # Test for downloading and installing a specific font
        It 'Downloads and installs all Nerd Fonts' {
            Should -Invoke -CommandName Install-Font -Exactly -Times 1
        }
    }
}