##########################################################################################################
#                                                                                                        #
#            _____                    _____                    _____                    _____            #
#           /\    \                  /\    \                  /\    \                  /\    \           #
#          /::\    \                /::\    \                /::\____\                /::\    \          #
#         /::::\    \               \:::\    \              /::::|   |               /::::\    \         #
#        /::::::\    \               \:::\    \            /:::::|   |              /::::::\    \        #
#       /:::/\:::\    \               \:::\    \          /::::::|   |             /:::/\:::\    \       #
#      /:::/__\:::\    \               \:::\    \        /:::/|::|   |            /:::/__\:::\    \      #
#     /::::\   \:::\    \              /::::\    \      /:::/ |::|   |           /::::\   \:::\    \     #
#    /::::::\   \:::\    \    ____    /::::::\    \    /:::/  |::|   | _____    /::::::\   \:::\    \    #
#   /:::/\:::\   \:::\____\  /\   \  /:::/\:::\    \  /:::/   |::|   |/\    \  /:::/\:::\   \:::\    \   #
#  /:::/  \:::\   \:::|    |/::\   \/:::/  \:::\____\/:: /    |::|   /::\____\/:::/  \:::\   \:::\____\  #
#  \::/   |::::\  /:::|____|\:::\  /:::/    \::/    /\::/    /|::|  /:::/    /\::/    \:::\  /:::/    /  #
#   \/____|:::::\/:::/    /  \:::\/:::/    / \/____/  \/____/ |::| /:::/    /  \/____/ \:::\/:::/    /   #
#         |:::::::::/    /    \::::::/    /                   |::|/:::/    /            \::::::/    /    #
#         |::|\::::/    /      \::::/____/                    |::::::/    /              \::::/    /     #
#         |::| \::/____/        \:::\    \                    |:::::/    /               /:::/    /      #
#         |::|  ~|               \:::\    \                   |::::/    /               /:::/    /       #
#         |::|   |                \:::\    \                  /:::/    /               /:::/    /        #
#         \::|   |                 \:::\____\                /:::/    /               /:::/    /         #
#          \:|   |                  \::/    /                \::/    /                \::/    /          #
#           \|___|                   \/____/                  \/____/                  \/____/           #
#                                                                                                        #
#                                                                                                        #
#     Author: David Pi                                                                                   #
#     Assisted by: ChatGPT - OpenAI                                                                      #
#     To Rina                                                                                            #
#                                                                                                        #
#     An automated backup script.                                                                        #
#                                                                                                        #
##########################################################################################################

$LogFile = Join-Path $PSScriptRoot "${env:COMPUTERNAME}.log"
$RoboCode = @{
    0 = 'No files were copied. No failure was encountered. No files were mismatched.'
    1 = 'All files were copied successfully.'
    2 = 'Additional files in the destination directory. No files were copied.'
    3 = 'Some files were copied. Additional files were present. No failure was encountered.'
    5 = 'Some files were copied. Some files were mismatched. No failure was encountered.'
    6 = 'Additional files and mismatched files exist. No files were copied and no failures were encountered.'
    7 = 'Files were copied, a file mismatch was present, and additional files were present.'
    8 = "Several files didn't copy."
}

function Read-Configuration ([string]$ConfigFile) {
    try {
        return Import-PowerShellDataFile -LiteralPath $ConfigFile -ErrorAction Stop
    }
    catch [System.Management.Automation.ItemNotFoundException] {
        $default = @'
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
'@
        Set-Content -LiteralPath $ConfigFile -Value $default
        Write-Host "An empty configuration file has been created at '${ConfigFile}'. Edit the file before running this script again."
        exit 1
    }
}

# Write log messages with a timestamp
function Write-Log ([string]$Message) {
    Add-Content -LiteralPath $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] ${Message}"
    Write-Host $Message
}

# Checks whether a backup is enabled and should run based on the configuration.
function Test-BackupEnabled ([hashtable]$config) {
    if ($config.Enabled) {
        if ($config.OnlyRunOn) {
            return ($config.OnlyRunOn -contains (Get-Date).DayOfWeek.ToString())
        }
        return $true
    }
    return $false
}

# Recursively expand environment variables
function Expand-EnvironmentVariables ([object]$InputObject) {
    if ($InputObject -is [string]) {
        return [System.Environment]::ExpandEnvironmentVariables($InputObject)
    }
    elseif ($InputObject -is [hashtable]) {
        foreach ($i in @($InputObject.Keys)) {
            $InputObject[$i] = Expand-EnvironmentVariables $InputObject[$i]
        }
    }
    elseif ($InputObject -is [array]) {
        $n = $InputObject.Length
        for ($i = 0; $i -lt $n; $i++) {
            $InputObject[$i] = Expand-EnvironmentVariables $InputObject[$i]
        }
    }
    return $InputObject
}

# Compress archive using 7-Zip
function Update-Archive ([hashtable]$config) {
    $config = Expand-EnvironmentVariables $config
    # Check for running processes in source directories
    if ($config.CheckProc -and (Test-ProcessPath -Path $config.Sources)) {
        Write-Log 'Skipping archiving: running processes in source directories.'
        return
    }
    # Define common switches for the 7-Zip command
    $switches = [System.Collections.Generic.List[string]]@(
        'u', '-up0q0r2x2y2z1w2', '-t7z', '-mx=9', '-ms=64m', '-mhe'
    )
    # Handle password encryption and add to switches
    if (-not [string]::IsNullOrEmpty($config.Password)) {
        try {
            $string = ConvertTo-SecureString -String $config.Password -ErrorAction Stop
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($string)
            $switches.Add("-p$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))")
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
        catch {
            Write-Log "Skipping archiving: password decryption failure. $($_.Exception.Message)"
            return
        }
    }
    # Add exclusion switches
    foreach ($s in $config.Exclusion) { $switches.Add("-xr!${s}") }
    # Execute 7-Zip with the defined switches and paths
    & $config.Executable $switches -- $config.Destination $config.Sources
    Write-Log "7-Zip finished with exit code ${LASTEXITCODE}."
}

# Tests if any running process paths match or are contained within a given set
# of paths. Constructs a trie structure and checks the paths against the trie.
function Test-ProcessPath ([string[]]$Path) {
    $sep = [IO.Path]::DirectorySeparatorChar
    $trie = @{}
    # Build trie
    foreach ($p in $Path) {
        try {
            $p = Resolve-Path -LiteralPath $p -ErrorAction Stop
        }
        catch {
            continue
        }
        $node = $trie
        foreach ($s in $p.Path.Split($sep)) {
            if (-not $s) { continue }
            if (-not $node.ContainsKey($s)) {
                $node[$s] = @{}
            }
            $node = $node[$s]
        }
        # Use $sep to mark the end of a path
        $node[$sep] = $true
    }
    # Check process paths against trie
    foreach ($p in (Get-Process | Select-Object -ExpandProperty Path -Unique)) {
        $node = $trie
        foreach ($s in $p.Split($sep)) {
            if (-not $s) { continue }
            if (-not $node.ContainsKey($s)) {
                break
            }
            $node = $node[$s]
            if ($node.ContainsKey($sep)) {
                return $true
            }
        }
    }
    return $false
}

# Backup OneDrive folder
function Backup-OneDrive ([hashtable]$config) {
    $config = Expand-EnvironmentVariables $config
    foreach ($d in $config.Source, $config.Destination) {
        if (-not $(try { Test-Path -LiteralPath $d -PathType Container } catch { $false })) {
            Write-Log "Skipping OneDrive backup: '${d}' is unreachable."
            return
        }
    }
    # Perform mirror copy with robocopy.
    robocopy $config.Source $config.Destination /MIR /DCOPY:DAT /J /COMPRESS /R:3 /MT /XA:S
    Write-Log "Robocopy finished backing up OneDrive with exit code ${LASTEXITCODE} ($($RoboCode[$LASTEXITCODE]))"
    # Apply unpinning to source path.
    if ($config.AutoUnpin) {
        Set-UnpinIfNotPinned -Path $config.Source
    }
}

<#
    Unpins all files in the OneDrive folder except those manually pinned (Always
    keeps on this device). Converts 'locally available' files to 'online-only'.

    - OneDrive file attributes:
        FILE_ATTRIBUTE_PINNED   : 0x00080000
        FILE_ATTRIBUTE_UNPINNED : 0x00100000

    - online-only
        FILE_ATTRIBUTE_PINNED   : False
        FILE_ATTRIBUTE_UNPINNED : True
        FILE_ATTRIBUTE_OFFLINE  : True

    - locally available
        FILE_ATTRIBUTE_PINNED   : False
        FILE_ATTRIBUTE_UNPINNED : False
        FILE_ATTRIBUTE_OFFLINE  : False

    - always available
        FILE_ATTRIBUTE_PINNED   : True
        FILE_ATTRIBUTE_UNPINNED : False
        FILE_ATTRIBUTE_OFFLINE  : False
#>
function Set-UnpinIfNotPinned ([string]$Path) {
    $PINNED = 0x00080000
    $UNPINNED = 0x00100000
    $COMBINED = $UNPINNED -bor $PINNED
    # If neither Unpinned nor Pinned, add Unpinned attribute.
    Get-ChildItem -LiteralPath $Path -Recurse -Attributes !Offline+!ReadOnly | ForEach-Object {
        if (!($_.Attributes -band $COMBINED)) {
            $_.Attributes = $_.Attributes -bor $UNPINNED
        }
    }
}

# Backup VMWare virtual machines
function Backup-VMWare ([hashtable]$config) {
    $config = Expand-EnvironmentVariables $config
    try {
        $rightDirs = @(Get-ChildItem -LiteralPath $config.Destination -Directory -ErrorAction Stop)
        $leftDirs = @(Get-ChildItem -LiteralPath $config.Source -Directory -ErrorAction Stop)
    }
    catch {
        Write-Log "Skipping VM backup: $($_.Exception.Message)"
        return
    }
    $exclusion = [System.Collections.Generic.List[string]]::new()
    # Exclude running VMs in the source.
    if ($config.SkipRunning) {
        $exts = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@('.lck', '.vmem', '.vmss'),
            [System.StringComparer]::OrdinalIgnoreCase
        )
        foreach ($dir in $leftDirs) {
            foreach ($file in Get-ChildItem -LiteralPath $dir.FullName) {
                if ($exts.Contains($file.Extension)) {
                    $exclusion.Add($dir.FullName)
                    Write-Log "Skipping '$($dir.Name)': VM not in a shutdown state."
                    break
                }
            }
        }
    }
    # Keep VMs that only exist in the destination. Excluding dirs on the right
    # prevents them from being removed.
    if ($config.KeepExtra) {
        if ($leftDirs.Length -eq $exclusion.Count) {
            # Nothing to backup, no need to proceed
            return
        }
        foreach ($dir in Compare-Object $leftDirs $rightDirs -Property Name -PassThru) {
            if ($dir.SideIndicator -eq '=>') {
                $exclusion.Add($dir.FullName)
            }
        }
    }
    # Mirror the directory using robocopy
    robocopy $config.Source $config.Destination /MIR /DCOPY:DAT /J /COMPRESS /R:3 /XF '*.log' '*.scoreboard' /XD 'caches' $exclusion
    Write-Log "Robocopy finished backing up VMs with exit code ${LASTEXITCODE} ($($RoboCode[$LASTEXITCODE]))"
}

# Import Configuration
$Configuration = Read-Configuration (Join-Path $PSScriptRoot "Config_${env:COMPUTERNAME}.psd1")

# Execute Archiving
if (Test-BackupEnabled -config $Configuration.Archive) {
    Update-Archive -config $Configuration.Archive
}

# Execute OneDrive Backup
if (Test-BackupEnabled -config $Configuration.OneDrive) {
    Backup-OneDrive -config $Configuration.OneDrive
}

# Execute VMWare Backup
if (Test-BackupEnabled -config $Configuration.VMWare) {
    Backup-VMWare -config $Configuration.VMWare
}
