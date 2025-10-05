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
#     Assisted by: ChatGPT                                                                               #
#     To Rina                                                                                            #
#                                                                                                        #
#     An automated backup script.                                                                        #
#                                                                                                        #
##########################################################################################################

param(
    [switch]$NoSkipping
)

$LogFile = Join-Path $PSScriptRoot "${env:COMPUTERNAME}.log"
$RoboParams = @('/MIR', '/J', '/DCOPY:DAT', '/R:3')
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
    $ConfigFile = Join-Path $PSScriptRoot "Config_${env:COMPUTERNAME}.psd1"
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
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')]
        [string]$Level = 'INFO'
    )

    $Message = "${Level}: ${Message}"
    Write-Host $Message
    Add-Content -LiteralPath $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] ${Message}"
}

# Checks whether a backup task is enabled and meets the conditions to run.
function Test-TaskCondition {
    param (
        [parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    if (-not $Config.Enable) { return $false }
    if ($NoSkipping) { return $true }
    if ($Config.DaysOfWeek -and (Get-Date).DayOfWeek.ToString() -notin $Config.DaysOfWeek) {
        return $false
    }
    if ($Config.NetworkName -and $Config.NetworkName -notin (Get-NetConnectionProfile).Name) {
        return $false
    }
    return $true
}

# Recursively expand environment variables
function Expand-EnvironmentVariables ([object]$InputObject) {
    if ($InputObject -is [string]) {
        return [System.Environment]::ExpandEnvironmentVariables($InputObject)
    }
    elseif ($InputObject -is [array]) {
        $n = $InputObject.Count
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

# Tests if any running processes are located within the specified paths.
function Test-RunningProcess {
    param([Parameter(Mandatory)][string[]]$Path)

    $Path = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue).Path
    if (-not $Path) { return $false }
    $Proc = Get-Process | Select-Object -ExpandProperty Path -Unique

    # Single path optimization
    if ($Path.Count -eq 1) {
        $p = $Path[0]
        # Path is a file
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            return $Proc.Contains($p)
        }
        # Add trailing slash to directory
        $p = Join-Path $p ''
        foreach ($s in $Proc) {
            if ($s.StartsWith($p)) { return $true }
        }
        return $false
    }

    # Build trie for multiple paths
    $sep = [IO.Path]::DirectorySeparatorChar
    $trie = @{}
    foreach ($p in $Path) {
        $node = $trie
        foreach ($s in $p.Split($sep)) {
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
    foreach ($p in $Proc) {
        $node = $trie
        foreach ($s in $p.Split($sep)) {
            if (-not $s) { continue }
            if (-not $node.ContainsKey($s)) { break }
            $node = $node[$s]
            if ($node.ContainsKey($sep)) {
                return $true
            }
        }
    }
    return $false
}

# Validates accessibility of source and destination directories. Returns them as
# strings.
function Test-SrcAndDst {
    param (
        [parameter(Mandatory)][string]$Source,
        [parameter(Mandatory)][string]$Destination
    )
    if ($Source -eq $Destination) {
        throw 'Source cannot be the same as Destination.'
    }
    foreach ($d in $Source, $Destination) {
        if (-not (Test-Path -LiteralPath $d -PathType Container)) {
            throw "'${d}' does not exist or is not a directory."
        }
    }
    return $Source, $Destination
}

function Test-PathModifiedSince {
    param(
        [Parameter(Mandatory)][string[]] $Path,
        [Parameter(Mandatory)][string] $RefFile
    )

    try {
        $RefTimeUtc = (Get-Item -LiteralPath $RefFile -ErrorAction Stop).LastWriteTimeUtc
    }
    catch {
        return $true
    }
    foreach ($p in $Path) {
        try { $p = Get-Item -LiteralPath $p -ErrorAction Stop } catch { continue }
        if ($p.LastWriteTimeUtc -gt $RefTimeUtc) { return $true }
        if (-not $p.PSIsContainer) { continue }

        if (Get-ChildItem -LiteralPath $p.FullName -Recurse -Attributes !ReparsePoint -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTimeUtc -gt $refTimeUtc } |
            Select-Object -First 1) {
            return $true
        }
    }
    return $false
}

# Update archives using 7-Zip
function Update-Archive {
    param (
        [parameter(Mandatory)][hashtable[]]$Configs
    )

    foreach ($Config in $Configs) {
        if (-not (Test-TaskCondition $Config)) { continue }
        $Config = Expand-EnvironmentVariables $Config

        [string[]]$srcs = $Config.Sources
        [string]$dst = $Config.Destination

        try {
            # Ensure each source currently exists.
            foreach ($s in $srcs) {
                if (-not (Test-Path -LiteralPath $s)) { throw "'$s' does not exist." }
            }

            # Ensure destination directory exists (create it if not)
            $s = Split-Path -Parent $dst
            if (-not (Test-Path -LiteralPath $s)) {
                New-Item -ItemType Directory -Path $s -Force -ErrorAction Stop | Out-Null
            }

            # Optional: block if processes are using sources
            if ($Config.CheckProc -and (Test-RunningProcess -Path $srcs)) {
                Write-Log "Skipping archiving: running processes in source directories. Destination: '$dst'." -Level WARNING
                continue
            }

            # Build switch list
            $switches = [System.Collections.Generic.List[string]]@('u', '-up0q0r2x2y2z1w2', '-t7z')
            $switches.AddRange([string[]]$Config.Parameters)

            # Only archive if source files are newer than the destination archive
            if ($Config.OnlyIfNewer) {
                if (-not (Test-PathModifiedSince -Path $srcs -RefFile $dst)) {
                    Write-Log "Skipping archiving: no newer source files. Destination: '$dst'." -Level WARNING
                    continue
                }
                $switches.Add('-stl')
            }

            # Password
            if (-not [string]::IsNullOrEmpty($Config.Password)) {
                $s = ConvertTo-SecureString -String $Config.Password -ErrorAction Stop
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
                $switches.Add('-mhe=on')
                $switches.Add("-p$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))")
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            }

            # Exclusions
            foreach ($s in @($Config.Exclusion)) { if ($s) { $switches.Add("-xr!${s}") } }

            # Execute 7-Zip
            & $Config.SevenZip $switches -- $dst $srcs
            $code = $LASTEXITCODE

            # 0: OK, 1: warning (non-fatal)
            if ($code -gt 1) { throw "7-Zip exited with code $code." }
            Write-Log ("Archiving finished. Exit code {0}. Destination: '{1}'. Size: {2} bytes." -f $code, $dst, (Get-Item -LiteralPath $dst).Length) -Level INFO
        }
        catch {
            Write-Log "Archiving failed. Destination: '${dst}'. $($_.Exception.Message)" -Level ERROR
        }
    }
}

# Backup general directories
function Backup-Directory {
    param (
        [parameter(Mandatory)][hashtable[]]$Configs
    )

    foreach ($Config in $Configs) {
        if (-not (Test-TaskCondition $Config)) { continue }
        $Config = Expand-EnvironmentVariables $Config
        try {
            # Check for accessibility
            $src, $dst = Test-SrcAndDst -Source $Config.Source -Destination $Config.Destination
            # Check for running processes
            if ($Config.CheckProc -and (Test-RunningProcess -Path $src)) {
                Write-Log "Skipping backup: running processes in the source directory. Source: '${src}'. Destination: '${dst}'." -Level WARNING
                continue
            }
            robocopy $src $dst $RoboParams $Config.RoboArgs
            Write-Log "Directory backup finished with exit code ${LASTEXITCODE}. ($($RoboCode[$LASTEXITCODE])) Source: '${src}'. Destination: '${dst}'." -Level INFO
        }
        catch {
            Write-Log "Directory backup failed. Source: '$($Config.Source)'. Destination: '$($Config.Destination)'. $($_.Exception.Message)" -Level ERROR
        }
    }
}

# Backup OneDrive folder
function Backup-OneDrive {
    param (
        [parameter(Mandatory)][hashtable]$Config
    )

    if (-not (Test-TaskCondition $Config)) { return }
    $Config = Expand-EnvironmentVariables $Config
    try {
        # Check for accessibility
        $src, $dst = Test-SrcAndDst -Source $Config.Source -Destination $Config.Destination
        # Check for running processes
        if ($Config.CheckProc -and (Test-RunningProcess -Path $src)) {
            Write-Log "Skipping backup: running processes in the source directory. Source: '${src}'. Destination: '${dst}'." -Level WARNING
            return
        }
        robocopy $src $dst $RoboParams /MT /XA:S
        Write-Log "OneDrive backup finished with exit code ${LASTEXITCODE}. ($($RoboCode[$LASTEXITCODE])) Source: '${src}'. Destination: '${dst}'." -Level INFO
    }
    catch {
        Write-Log "OneDrive backup failed. Source: '$($Config.Source)'. Destination: '$($Config.Destination)'. $($_.Exception.Message)" -Level ERROR
        return
    }
    # Apply unpinning to the source path.
    if ($Config.AutoUnpin) {
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
        [parameter(Mandatory)][string]$Path
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

# Backup VMware virtual machines
function Backup-VMware {
    param (
        [parameter(Mandatory)][hashtable]$Config
    )

    if (-not (Test-TaskCondition $Config)) { return }
    $Config = Expand-EnvironmentVariables $Config
    try {
        # Check for accessibility
        $src, $dst = Test-SrcAndDst -Source $Config.Source -Destination $Config.Destination
        $leftDirs = @(Get-ChildItem -LiteralPath $src -Directory -ErrorAction Stop)
        $rightDirs = @(Get-ChildItem -LiteralPath $dst -Directory -ErrorAction Stop)
        # Exclude running VMs in the source.
        $exclusion = [System.Collections.Generic.List[string]]::new()
        if ($Config.SkipRunning) {
            $exts = [System.Collections.Generic.HashSet[string]]::new(
                [string[]]@('.lck', '.vmem', '.vmss'),
                [System.StringComparer]::OrdinalIgnoreCase
            )
            foreach ($dir in $leftDirs) {
                foreach ($file in Get-ChildItem -LiteralPath $dir.FullName -ErrorAction SilentlyContinue) {
                    if ($exts.Contains($file.Extension)) {
                        $exclusion.Add($dir.FullName)
                        Write-Log "Skipping running VM: '$($dir.Name)'." -Level WARNING
                        break
                    }
                }
            }
        }
        # Keep VMs that only exist in the destination. Excluding dirs on the right
        # prevents them from being removed.
        if ($Config.KeepExtra) {
            if ($leftDirs.Count -eq $exclusion.Count) {
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
        Write-Log "VMware backup finished with exit code ${LASTEXITCODE}. ($($RoboCode[$LASTEXITCODE])) Source: '${src}'. Destination: '${dst}'." -Level INFO
    }
    catch {
        Write-Log "VMware backup failed: Source: '$($Config.Source)'. Destination: '$($Config.Destination)'. $($_.Exception.Message)" -Level ERROR
    }
}

# Import Configuration
$Configuration = Read-Configuration

# Execute Backups
Update-Archive $Configuration.Archives
Backup-Directory $Configuration.Directories
Backup-OneDrive $Configuration.OneDrive
Backup-VMware $Configuration.VMware
