# Windows-Explorer-Backup

PowerShell scripts to back up and re-open a Windows Explorer session (all opened folders).

## Requirements

- [PowerShell 7+](https://github.com/PowerShell/PowerShell)

## Usage

### Backing Up Sessions

1. Create a folder to store the Explorer session text files.
2. To back up the current session, call `Backup-Session.ps1 -backupsPath [folderPath]`, where
   `[folderPath]` is the full path to that folder.

### Restoring Sessions

- To restore the most recently backed up Explorer session, call
  `Restore-Session.ps1 -backupsPath [folderPath]`.
- To restore a specific Explorer session, call 
  `Restore-Session.ps1 -backupsPath [folderPath] -sessionFile [fileName]`, where `[fileName]` is the
  name (not full path) of the session file to restore from.

## Current Limitations

### Explorer Windows Sizes and Positions

The sizes and positions of Windows Explorer windows are not backed up, and thus also not restored.  
Windows are opened with whatever is currently the default size and position for Windows Explorer.

### Multi-Desktop Support

There is currently no support for multiple/virtual desktops. The backup process will back up all
windows across all desktops, and the restore process will restore all windows on the current desktop
(I think, I have not done any testing).

### Session Management

- There is currently no form of session management.
- Sessions can not be mixed or merged.  
- Restoring a session on top of already opened Windows Explorer windows will simply add the restored
  windows to the already existing ones.
- There is no checking for duplicate windows (the same folder/path being open in multiple Explorer
  windows). If a specific folder/path is already open in Windows Explorer, it will simply be opened
  again in another Explorer window.

### Scheduling/Auto-Backups

Currently, backups and restores can only be triggered manually. There is no support for automatic
backups or restores.  
Scheduling these can be set up manually via the Windows Task Scheduler.

## Performance Considerations

- Each individual Explorer window is opened as an entirely separate `explorer.exe` process. This was
  by design to improve robustness against crashes and freezes of Explorer windows.
  - This does lead to increased system memory usage if many Explorer windows are restored.
  - Currently, there is no setting to change this behavior. 
  
## Other Considerations

### Unicode Support

The session files are stored as `UTF-8`. Folder paths with non-`ASCII` characters *should* work. In
theory, all characters that Windows allows in folder paths should work. If you have issues, please
open an issue!
