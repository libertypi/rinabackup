# RinaBackup

RinaBackup is an automated backup script tailored for my personal use. Written in PowerShell, it is primarily designed to backup files from PCs and laptops to a centralized storage. After configuration, it is recommended to set the script up to run automatically via Task Scheduler.

- Incrementally backs up files into 7-Zip archives or directories.
- Special handling for OneDrive folders and VMware VMs.
- Comprehensive and configurable backup controls.

#### Configuration File

Upon the first run, the script generates a configuration file `Config_%COMPUTERNAME%.psd1` in its directory. Edit the file before running this script again.

```powershell
<#
    Common Options:
    - Enable:       Enable or disable the backup section.
    - DaysOfWeek:   Specifies days of the week for backup; an empty value skips the check.
                    Accepts: @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday').
    - NetworkName:  Run only if connected to the specific network; an empty value skips the check.
                    Get network name with cmdlet: `Get-NetConnectionProfile`
    - CheckProc:    Skip action if running processes are found in source.
    - OnlyIfNewer:  Scan Sources and run 7-Zip only if any file is newer than the existing archive.
                    Archive time will be set to latest file time.

    Archives Options:
    - SevenZip:     Path to the 7-Zip executable.
    - Sources:      Source directories to archive.
    - Destination:  Destination path for the archive.
    - Exclusion:    Patterns to exclude from archiving.
    - Password:     Password for archive (DPAPI encrypted, see Notes).
    - Parameters:   Additional parameters passed to 7-Zip, e.g.: @('-mx=9', '-ms=32m')

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
      They can be created with the following command:
        Read-Host -AsSecureString | ConvertFrom-SecureString
#>

@{
    Archives    = @(
        @{
            Enable      = $false
            DaysOfWeek  = @()
            NetworkName = ''
            CheckProc   = $true
            OnlyIfNewer = $false
            SevenZip    = ''
            Sources     = @()
            Destination = ''
            Exclusion   = @()
            Password    = ''
            Parameters  = @()
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
```
