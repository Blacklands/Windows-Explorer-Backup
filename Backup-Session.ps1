#requires -PSEdition Core

param (
    [Parameter(Mandatory=$true)][string]$backupsPath
)

if ((Test-Path $backupsPath -PathType Container) -eq $false) {
        # Ugly workaround because "Write-Error" formats really weirdly. :/
        [Console]::ForegroundColor = 'red'
        [Console]::Error.WriteLine("The path for the backup files provided, `"$backupsPath`", does not exist, or is not a path to a folder. Quitting.")
        [Console]::ResetColor()
        Exit 1
}

Write-Host "Collecting folders opened in the current Windows Explorer session..."

$sessionFolders = @((New-Object -com shell.application).Windows()).Document.Folder.Self.Path

Write-Host "The current Windows Explorer session consists of the following $($sessionFolders.Count) open folders:"
$sessionFolders | ForEach-Object { "- `"$_`"" } | Write-Host -Separator "\n"

$backupsFile = "$(Get-Date -Format 'yyyy-MM-ddTHH-mm-ss').txt"

Write-Host "Saving Explorer session backup file `"$backupsFile`" to path `"$backupsPath`"..."

$sessionFolders | Out-File -LiteralPath (Join-Path -Path $backupsPath -ChildPath $backupsFile) -Encoding utf8