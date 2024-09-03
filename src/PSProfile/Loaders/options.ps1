Set-PSReadLineOption -HistorySearchCursorMovesToEnd
Set-PSReadlineOption -BellStyle None #Disable ding on typing error
Set-PSReadlineOption -EditMode Emacs #Make TAB key show parameter options

Set-PSReadlineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadlineKeyHandler -Key DownArrow -Function HistorySearchForward 