#requires -PSEdition Core

param (
    [Parameter(Mandatory=$true)][string]$backupsPath,
    [Parameter(Mandatory=$false)][string]$sessionFile
)

function WriteErrorAndQuit {
    
    param (
        [string]$error
    )
    
    # Ugly workaround because "Write-Error" formats really weirdly. :/
    [Console]::ForegroundColor = 'red'
    [Console]::Error.WriteLine($error)
    [Console]::ResetColor()
    Exit 1
}

if ((Test-Path $backupsPath -PathType Container) -eq $false) {
        WriteErrorAndQuit("The path for the backup files provided, `"$backupsPath`", does not exist, or is not a path to a folder. Quitting.")
}

Write-Host "Restoring Windows Explorer session from backups path `"$backupsPath`"..."

$restoreFile = $sessionFile

if ($restoreFile -eq "") {
    
    Write-Host "No session backup file was given. Using the most recent existing one."
    
    $filesAll = Get-ChildItem -Path $backupsPath -Name -Recurse
    $filesText = $filesAll | Where-Object { $_ -like "*.txt" }
    
    if ($filesText.Count -eq 0) {
        WriteErrorAndQuit("No session backups were found under the given path. Quitting.")  
    }
    
    $filesSorted = $filesText | Sort-Object -Descending
    
    $restoreFile = $filesSorted[0]
}
else {
    
    if (-Not ($restoreFile -like "*.txt")) {
        WriteErrorAndQuit("The given file, `"$restoreFile`", does not seem to be a session backup file, which should be a text file. Quitting.")     
    }
}

$restoreFilePath = Join-Path -Path $backupsPath -ChildPath $restoreFile

if ((Test-Path $restoreFilePath -PathType Leaf) -eq $false) {
        WriteErrorAndQuit("The given file name, `"$restoreFile`", does not seem to point to an actual file. Quitting.")       
}

Write-Host "Restoring from session file `"$restoreFile`"..."

$folders = Get-Content -Path $restoreFilePath

if ($folders.Count -eq 0) {
    
    Write-Host "There are no folders to restore."
    Exit 0
}

Write-Host "Restoring Windows Explorer session consisting of the following $($folders.Count) folders:"
$folders | ForEach-Object { "- `"$_`"" } | Write-Host -Separator "\n"

foreach ($folder in $folders) {
    Write-Host "Opening folder `"$folder`"..."
    # Need "-LiteralPath" here or "functional characters" like "[" and "]" break "Invoke-Item".
    Invoke-Item -LiteralPath $folder
    # If we open folders too quickly at once, Windows Explorer can apparently crash. So let's give it more time.
    Start-Sleep -Milliseconds 100
}

Write-Host "Successfully restored Explorer session."