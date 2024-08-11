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