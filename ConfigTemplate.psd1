<#
    Common Options:
    - Enable:       Enable or disable the backup section.
    - DaysOfWeek:   Specifies days of the week for backup; an empty value skips the check.
                    Accepts: @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday').
    - NetworkName:  Run only if connected to the specific network; an empty value skips the check.
                    Get network name with cmdlet: `Get-NetConnectionProfile`
    - CheckProc:    Skip action if running processes are found in source.

    Archives Options:
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
    Archives    = @(
        @{
            Enable      = $false
            DaysOfWeek  = @()
            NetworkName = ''
            CheckProc   = $true
            Executable  = ''
            Sources     = @()
            Destination = ''
            Exclusion   = @()
            Password    = ''
        }
    )
    Directories = @(
        @{
            Enable      = $false
            DaysOfWeek  = @()
            NetworkName = ''
            CheckProc   = $false
            Source      = ''
            Destination = ''
            RoboArgs    = @()
        }
    )
    OneDrive    = @{
        Enable      = $false
        DaysOfWeek  = @()
        NetworkName = ''
        CheckProc   = $false
        Source      = '%OneDrive%'
        Destination = ''
        AutoUnpin   = $false
    }
    VMWare      = @{
        Enable      = $false
        DaysOfWeek  = @()
        NetworkName = ''
        SkipRunning = $true
        KeepExtra   = $true
        Source      = ''
        Destination = ''
    }
}