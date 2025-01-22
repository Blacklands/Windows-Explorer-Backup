#requires -PSEdition Core

using module ".\Set-WindowState.psm1"

param (
    [Parameter(Mandatory=$true)][string]$backupsPath,
    [Parameter(Mandatory=$false)][string]$sessionFile
)

function WriteError {
    
    param (
        [string]$error
    )
    
    # Ugly workaround because "Write-Error" formats really weirdly. :/
    [Console]::ForegroundColor = 'red'
    [Console]::Error.WriteLine($error)
    [Console]::ResetColor()
}

function WriteErrorAndQuit {
    
    param (
        [string]$error
    )
    
    WriteError $error
    Exit 1
}

if ((Test-Path -LiteralPath $backupsPath -PathType Container) -eq $false) {
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

if ((Test-Path -LiteralPath $restoreFilePath -PathType Leaf) -eq $false) {
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
    
    # If we try to open an invalid folder path with Explorer, it won't skip it, it will open a generic folder (apparently the
    # "Documents" folder of the user) instead, which we don't want. So we need to manually check and skip those.
    if ((Test-Path -LiteralPath $folder) -eq $false) {
        WriteError "The folder `"$folder`" does not seem to exist, or is not accessible. Skipping it."
        continue
    }
    
    Write-Host "Opening folder `"$folder`"..."
    # We're using "-Wait", also because things break if trying to launch Windows Explorer processess too fast while also using the
    # "Miminized" window style - the Explorer windows don't open correctly and/or are super tiny and render wrong.
    # "-Wait" is pretty slow. This is probably because we are actually starting an entirely separate process
    # for each window. This is actually good because it prevents all windows from crashing and closing if there is an issue with a
    # single one. This is supposed to be a toggleable option in Explorer anyway, but the option has been broken and does not work for
    # (I think) years at this point.
    # This likely uses a lot more system memory. But who cares? We have enough.
    $process = Start-Process -PassThru -Wait -WindowStyle Normal -FilePath explorer.exe -ArgumentList "$folder"
    # We're doing this instead of using the "Minimized" window style option above because if the windows are starting off minimized,
    # Windows does not render a taskbar preview for any of them (they will just be blank) until they have been brought to the foreground
    # once.
    # Opening them in the foreground for a moment and then minimizing them makes that work.
    # The disadvantage of this is that the "ALT + TAB" menu will get messed up, with all the windows at the top (since they were just in
    # focus), which does not happen if they are opened in a minimized state.
    # I have not been able to find a way to "refresh" the minimized windows without bringing them to the foreground, so I don't think
    # a "best of both worlds" solution exists - at least not without me diving deeper into the Windows API and seeing if there is some
    # low-level way to do this.
    # 
    # Also, we need to "generically" get all processes with the name "explorer" here, which is also not ideal. For some reason, the
    # process that "Start-Process" returns has the wrong ID (one that does not exist) and does not actually point to the Explorer
    # process that gets started - we have no idea what the ID of that process is, and we can't even find it by name, either.
    # Yes, this means that all Explorer windows will get minimized when this script is run, even the ones that were already open before.
    # I don't think we can do anything about that if we're using this method.
    Get-Process "explorer" | Set-WindowState -State MINIMIZE
    # Alternative way of giving Explorer more time between launches. Does NOT seem to consistently work properly with the "Minimized"
    # window style option, things can still break. Probably okay with the "Normal" option, but still iffy. Explorer really needs some
    # time between launches, so using the "-Wait" option makes more sense I think.
    # Start-Sleep -Milliseconds 200
}

Write-Host "Successfully restored Explorer session."