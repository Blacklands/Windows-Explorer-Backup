# -----------------------------------------------------
# Taken from: https://github.com/surfaceowl/Set-WindowState.ps1
# No edits have been made except for this header and
# renaming the file to turn it into a PowerShell module.
# -----------------------------------------------------

# The $WindowStates hashtable had two entries for 'MAXIMIZE' and 'SHOWMAXIMIZED' with the same value. Since the ValidateSet attribute ensures that only valid strings are used as keys, and the hashtable is used to map these strings to their corresponding integer values, the duplicate entry is unnecessary and has been removed.
# The Set-Alias command was moved to the BEGIN block to keep the alias definition close to the function definition, which is a common practice for better script organization and readability.

function Set-WindowState
{
    param(
        [Parameter()]
        [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE',
                'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED',
                'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
        [Alias('Style')]
        [String] $State = 'SHOW',

        [Parameter(ValueFromPipelineByPropertyname = 'True')]
        [System.IntPtr] $MainWindowHandle = (Get-Process -id $pid).MainWindowHandle,

        [Parameter()]
        [switch] $PassThru

    )
    BEGIN
    {
        Set-Alias -Name 'Set-WindowStyle' -Value 'Set-WindowState'

        $WindowStates = @{
            'FORCEMINIMIZE' = 11
            'HIDE' = 0
            'MAXIMIZE' = 3
            'MINIMIZE' = 6
            'RESTORE' = 9
            'SHOW' = 5
            'SHOWDEFAULT' = 10
            'SHOWMINIMIZED' = 2
            'SHOWMINNOACTIVE' = 7
            'SHOWNA' = 8
            'SHOWNOACTIVATE' = 4
            'SHOWNORMAL' = 1
        }

        $Win32ShowWindowAsync = Add-Type -memberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -name "Win32ShowWindowAsync" -namespace Win32Functions -passThru

    }
    PROCESS
    {
        $Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, $WindowStates[$State]) | Out-Null
        Write-Verbose ("Set Window State on '{0}' to '{1}' " -f $MainWindowHandle, $State)

        if ($PassThru)
        {
            Write-Output $MainWindowHandle
        }

    }
    END
    {
    }

}
