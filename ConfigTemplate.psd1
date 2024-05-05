<#
    Common Options:
    - Enabled:      Enable or disable the backup section.
    - OnlyRunOn:    Specifies days of the week for backup; an empty value skips the check.
                    Accepts: @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday').

    Archive Options:
    - CheckProc:    Skip archiving if running processes are found in source.
    - Executable:   Path to the 7-Zip executable.
    - Sources:      Source directories to archive.
    - Destination:  Destination path for the archive.
    - Exclusion:    Patterns to exclude from archiving.
    - Password:     Password for archive (DPAPI encrypted).

    Directories Options:
    - Source:       Source directory for backup.
    - Destination:  Destination path for backup.
    - RoboArgs:     Additional arguments passed to robocopy.

    OneDrive Options:
    - AutoUnpin:    Free up space except for "always keeps on this device" files.
    - Source:       Source directory for OneDrive backup.
    - Destination:  Destination path for OneDrive backup.

    VMWare Options:
    - SkipRunning:  Skip running VMs.
    - KeepExtra:    Keep extra VMs in destination not found in source.
    - Source:       Source directory for VMWare backup.
    - Destination:  Destination path for VMWare backup.

    Notes:
    - Environment variables can be used, enclosed in '%' (e.g. '%USERPROFILE%').
    - Passwords require an encrypted string using the Windows Data Protection API (DPAPI).
      They can be created with the following command: `Read-Host -AsSecureString | ConvertFrom-SecureString`
#>

@{
    Archive     = @{
        Enabled     = $false
        OnlyRunOn   = @()
        CheckProc   = $true
        Executable  = ''
        Sources     = @()
        Destination = ''
        Exclusion   = @()
        Password    = ''
    }
    Directories = @(
        @{
            Enabled     = $false
            OnlyRunOn   = @()
            CheckProc   = $false
            Source      = ''
            Destination = ''
            RoboArgs    = @()
        }
    )
    OneDrive    = @{
        Enabled     = $false
        OnlyRunOn   = @()
        AutoUnpin   = $false
        Source      = '%OneDrive%'
        Destination = ''
    }
    VMWare      = @{
        Enabled     = $false
        OnlyRunOn   = @()
        SkipRunning = $true
        KeepExtra   = $true
        Source      = ''
        Destination = ''
    }
}