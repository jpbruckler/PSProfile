# Add PSReadline settings and keybindings here

Set-PSReadLineOption -HistorySearchCursorMovesToEnd
Set-PSReadlineOption -BellStyle None #Disable ding on typing error
Set-PSReadlineOption -EditMode Emacs #Make TAB key show parameter options

Set-PSReadlineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadlineKeyHandler -Key DownArrow -Function HistorySearchForward

$GitCommitScript = {
    # Define commit types
    $commitTypes = @("feat", "fix", "docs", "style", "refactor", "perf", "test", "chore", "build", "ci", "revert")
    $menuItems = [System.Collections.Generic.List[string]]::new()
    $commitTypes | ForEach-Object { $menuItems.Add($_) }

    # Prompt for commit type in the console
    $commitType = New-InteractiveMenu -menuItems $menuItems -title "Select commit type" -hint "Use arrows or type the number. 'Enter' - Select, 'ESC' - Exit"

    # Prompt for optional scope
    $scope = Read-Host "Enter scope (optional, e.g., parser):"

    # Check for breaking change
    $isBreakingChange = Read-Host "Is this a breaking change? (y/n):"
    $breakingIndicator = ""
    if ($isBreakingChange -eq "y") {
        $breakingIndicator = "!"
    }

    # Prompt for commit message
    $commitMessage = Read-Host "Enter commit message:"

    if (-not $commitMessage) {
        Write-Host "Commit aborted: No commit message entered." -ForegroundColor Red
        return
    }

    # Construct the full commit message using the -f format operator
    $scopePart = if ($scope) { "({0})" -f $scope } else { "" }
    $fullCommitMessage = "{0}{1}{2}: {3}" -f $commitType, $scopePart, $breakingIndicator, $commitMessage

    # Add body or footer
    $additionalDetails = Read-Host "Add a detailed description or footer (leave blank to skip):"

    # Commit and push
    Write-Host 'Committing and pushing changes...' -ForegroundColor Yellow
    git add .
    git commit -m "`"$fullCommitMessage`"" -m "`"$additionalDetails`""
    git push origin HEAD

    Write-Host "Commit and push completed successfully." -ForegroundColor Green
    return
}

# Register PSReadline key binding for the scriptblock
Set-PSReadlineKeyHandler -Key Ctrl+g -ScriptBlock $GitCommitScript