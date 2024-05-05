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
$RoboParams = @('/MIR', '/J', '/MT', '/DCOPY:DAT', '/R:3')
$RoboCode = @{
    0 = 'No files copied or mismatched. No failure.'
    1 = 'All files copied.'
    2 = 'Extra files in destination. No copy.'
    3 = 'Some files copied. Extra files. No failure.'
    5 = 'Some files copied or mismatched. No failure.'
    6 = 'Extra and mismatched files. No copy or failure.'
    7 = 'Files copied with mismatches and extras.'
    8 = 'Some files not copied.'
}

function Read-Configuration {
    param (
        [parameter(Mandatory = $true)]
        [string]$ConfigFile
    )

    try {
        return Import-PowerShellDataFile -LiteralPath $ConfigFile -ErrorAction Stop
    }
    catch [System.Management.Automation.ItemNotFoundException] {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ConfigTemplate.psd1') -Destination $ConfigFile
        Write-Host "An empty configuration file has been created at '${ConfigFile}'. Edit the file before running this script again."
        exit 1
    }
}

# Write log messages with a timestamp
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')]
        [string]$Level = 'INFO'
    )

    $Message = "${Level}: ${Message}"
    Write-Host $Message
    Add-Content -LiteralPath $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] ${Message}"
}

# Checks whether a backup is enabled and should run based on the configuration.
function Test-BackupEnabled([hashtable]$config) {
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
    elseif ($InputObject -is [array]) {
        $n = $InputObject.Length
        for ($i = 0; $i -lt $n; $i++) {
            $InputObject[$i] = Expand-EnvironmentVariables $InputObject[$i]
        }
    }
    elseif ($InputObject -is [hashtable]) {
        foreach ($i in @($InputObject.Keys)) {
            $InputObject[$i] = Expand-EnvironmentVariables $InputObject[$i]
        }
    }
    return $InputObject
}

# Tests if any running process paths match or are contained within a given set
# of paths. Constructs a trie structure and checks the paths against the trie.
function Test-ProcessPath {
    param (
        [parameter(Mandatory = $true)]
        [string[]]$Path
    )

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

# Validates accessibility of source directory and creates destination directory
# if it does not exist. Return source and destination as strings.
function Test-SrcAndDst {
    param (
        [parameter(Mandatory = $true)]
        [string]$Source,

        [parameter(Mandatory = $true)]
        [string]$Destination
    )
    if (-not ($Source -and $Destination)) {
        throw 'Path value cannot be empty or null.'
    }
    if ($Source -eq $Destination) {
        throw 'Source cannot be the same as Destination.'
    }
    if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
        throw "$Source is unreachable or not a directory."
    }
    if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
        New-Item -ItemType Directory -Path $Destination -ErrorAction Stop
    }
    return $Source, $Destination
}

# Compress archive using 7-Zip
function Update-Archive {
    param (
        [parameter(Mandatory = $true)]
        [hashtable]$config
    )

    if (-not (Test-BackupEnabled $config)) { return }
    $config = Expand-EnvironmentVariables $config
    [string[]]$srcs = $config.Sources
    [string]$dst = $config.Destination
    [string]$exe = $config.Executable

    # Check for accessibility of source files and 7zip executable
    foreach ($f in ($srcs + $exe)) {
        if (-not ($f -and (Test-Path -LiteralPath $f))) {
            Write-Log "Skipping archiving: ${f} is not accessible." -Level ERROR
            return
        }
    }
    # Check for running processes in source directories
    if ($config.CheckProc -and (Test-ProcessPath -Path $srcs)) {
        Write-Log 'Skipping archiving: running processes in source directories.' -Level WARNING
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
            Write-Log "Skipping archiving: password decryption failure. $($_.Exception.Message)" -Level ERROR
            return
        }
    }
    # Add exclusion switches
    foreach ($s in $config.Exclusion) { $switches.Add("-xr!${s}") }
    # Execute 7-Zip with the defined switches and paths
    & $exe $switches -- $dst $srcs
    Write-log "Archiving finished. Dst: '${dst}'. Exit code: ${LASTEXITCODE}. Size: $((Get-Item -LiteralPath $dst).Length)." -Level INFO
}

# Backup general directories
function Backup-Directory {
    param (
        [parameter(Mandatory = $true)]
        [hashtable[]]$configs
    )

    foreach ($config in $configs) {
        if (-not (Test-BackupEnabled $config)) { continue }
        $config = Expand-EnvironmentVariables $config

        # Check for accessibility
        try {
            $src, $dst = Test-SrcAndDst -Source $config.Source -Destination $config.Destination
        }
        catch {
            Write-Log "Skipping '$($config.Source)': $($_.Exception.Message)" -Level ERROR
            continue
        }
        # Check for running processes
        if ($config.CheckProc -and (Test-ProcessPath -Path $src)) {
            Write-Log "Skipping '${src}': running processes in directory." -Level WARNING
            continue
        }
        # Perform mirror copy with robocopy.
        robocopy $src $dst $RoboParams $config.RoboArgs
        Write-Log "Directory backup finished. Src: '${src}'. Dst: '${dst}'. Exit code: ${LASTEXITCODE} ($($RoboCode[$LASTEXITCODE]))" -Level INFO
    }
}

# Backup OneDrive folder
function Backup-OneDrive {
    param (
        [parameter(Mandatory = $true)]
        [hashtable]$config
    )

    if (-not (Test-BackupEnabled $config)) { return }
    $config = Expand-EnvironmentVariables $config

    # Check for accessibility
    try {
        $src, $dst = Test-SrcAndDst -Source $config.Source -Destination $config.Destination
    }
    catch {
        Write-Log "Skipping OneDrive backup: $($_.Exception.Message)" -Level ERROR
        return
    }
    # Perform mirror copy with robocopy.
    robocopy $src $dst $RoboParams /XA:S
    Write-Log "OneDrive backup finished. Src: '${src}'. Dst: '${dst}'. Exit code: ${LASTEXITCODE} ($($RoboCode[$LASTEXITCODE]))" -Level INFO
    # Apply unpinning to source path.
    if ($config.AutoUnpin) {
        Set-UnpinIfNotPinned -Path $src
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
function Set-UnpinIfNotPinned {
    param (
        [parameter(Mandatory = $true)]
        [string]$Path
    )

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
function Backup-VMWare {
    param (
        [parameter(Mandatory = $true)]
        [hashtable]$config
    )

    if (-not (Test-BackupEnabled $config)) { return }
    $config = Expand-EnvironmentVariables $config

    # Check for accessibility
    try {
        $src, $dst = Test-SrcAndDst -Source $config.Source -Destination $config.Destination
        $leftDirs = @(Get-ChildItem -LiteralPath $src -Directory -ErrorAction Stop)
        $rightDirs = @(Get-ChildItem -LiteralPath $dst -Directory -ErrorAction Stop)
    }
    catch {
        Write-Log "Skipping VMWare backup: $($_.Exception.Message)" -Level ERROR
        return
    }
    # Exclude running VMs in the source.
    $exclusion = [System.Collections.Generic.List[string]]::new()
    if ($config.SkipRunning) {
        $exts = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@('.lck', '.vmem', '.vmss'),
            [System.StringComparer]::OrdinalIgnoreCase
        )
        foreach ($dir in $leftDirs) {
            foreach ($file in Get-ChildItem -LiteralPath $dir.FullName) {
                if ($exts.Contains($file.Extension)) {
                    $exclusion.Add($dir.FullName)
                    Write-Log "Skipping '$($dir.Name)': VM not in a shutdown state." -Level WARNING
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
    robocopy $src $dst $RoboParams /XF '*.log' '*.scoreboard' /XD 'caches' $exclusion
    Write-Log "VMware backup finished. Src: '${src}'. Dst: '${dst}'. Exit code: ${LASTEXITCODE} ($($RoboCode[$LASTEXITCODE]))" -Level INFO
}

# Import Configuration
$Configuration = Read-Configuration (Join-Path $PSScriptRoot "Config_${env:COMPUTERNAME}.psd1")

# Execute Backups
Update-Archive $Configuration.Archive
Backup-Directory $Configuration.Directories
Backup-OneDrive $Configuration.OneDrive
Backup-VMWare $Configuration.VMWare
