# RinaBackup

RinaBackup is an automated backup script tailored for my personal use, featuring:

- Updating multiple folders into an encrypted 7-Zip archive, skips if running processes are found in the source, and securely stores the password.
- Backs up the OneDrive folder, freeing up local space after backup, except for files marked as "always keep on this device."
- Backs up VMware virtual machines to another location, skipping running VMs.

Upon the first run, the script generates a configuration file, `"Config_%COMPUTERNAME%.psd1"`, in its directory. This file offers more options to meet many twisted needs. After configuration, it’s recommended to set the script up to run automatically via the Task Scheduler.

#### Configuration Template

```powershell
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
    Archive  = @{
        Enabled     = $false
        OnlyRunOn   = @()
        CheckProc   = $true
        Executable  = ''
        Sources     = @()
        Destination = ''
        Exclusion   = @()
        Password    = ''
    }
    OneDrive = @{
        Enabled     = $false
        OnlyRunOn   = @()
        AutoUnpin   = $false
        Source      = '%OneDrive%'
        Destination = ''
    }
    VMWare   = @{
        Enabled     = $false
        OnlyRunOn   = @()
        SkipRunning = $true
        KeepExtra   = $true
        Source      = ''
        Destination = ''
    }
}
```
