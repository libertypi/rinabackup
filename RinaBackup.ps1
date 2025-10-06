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
    [switch]$NoSkipping,
    [switch]$Debug
)

$LogFile = Join-Path $PSScriptRoot "${env:COMPUTERNAME}.log"
$RoboParams = @('/MIR', '/J', '/DCOPY:DAT', '/R:3')
$RoboCode = @{
    0  = 'No changes.'
    1  = 'Copied successfully.'
    2  = 'No copy. Extras present.'
    3  = 'Some files copied. Extras present.'
    4  = 'Mismatches present.'
    5  = 'Some files copied. Mismatches present.'
    6  = 'No copy. Extras and mismatches present.'
    7  = 'Some files copied. Extras and mismatches present.'
    8  = 'Failure: some files/dirs not copied.'
    16 = 'Serious error: nothing copied.'
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
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')]
        [string]$Level = 'INFO'
    )

    $Message = "${Level}: ${Message}"
    $fg = switch ($Level) {
        'DEBUG' { 'DarkGray' }
        'INFO' { 'Gray' }
        'WARNING' { 'Yellow' }
        'ERROR' { 'Red' }
        'CRITICAL' { 'Red' }
    }
    Write-Host $Message -ForegroundColor $fg

    if ($Level -eq 'DEBUG' -and -not $Debug) { return }
    Add-Content -LiteralPath $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm')] ${Message}"
}

function ConvertTo-HumanSize {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [long]$Bytes
    )
    begin {
        $units = @('B', 'KiB', 'MiB', 'GiB', 'TiB', 'PiB', 'EiB', 'ZiB', 'YiB')
    }
    process {
        if ($Bytes -eq 0) { return '0.00 B' }
        $exp = [math]::Min(
            [math]::Floor([math]::Log([math]::Abs($Bytes), 1024)),
            $units.Count - 1
        )
        [string]::Format(
            [Globalization.CultureInfo]::InvariantCulture,
            '{0:0.00} {1}',
            $Bytes / [math]::Pow(1024, $exp),
            $units[$exp]
        )
    }
}

# Checks whether a backup task is enabled and meets the conditions to run.
function Test-TaskCondition {
    param (
        [parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    if (-not $Config.Enable) { return $false }
    if ($NoSkipping) { return $true }
    if ($Config.DaysOfWeek -and (Get-Date).DayOfWeek.ToString() -notin $Config.DaysOfWeek) { return $false }
    if ($Config.NetworkName -and $Config.NetworkName -notin (Get-NetConnectionProfile).Name) { return $false }
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

# Ensure source and destination directories are reachable. Returns them as strings.
function Test-SrcAndDestDirs {
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Source,

        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Destination
    )

    if ($Source -ieq $Destination) {
        throw 'Source and Destination cannot be the same.'
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
        [Parameter(Mandatory)][string[]]$Path,
        [Parameter(Mandatory)][string]$RefFile
    )

    try { $RefTimeUtc = (Get-Item -LiteralPath $RefFile -ErrorAction Stop).LastWriteTimeUtc }
    catch { return $true }

    foreach ($p in $Path) {
        try { $p = Get-Item -LiteralPath $p -ErrorAction Stop } catch { continue }
        if ($p.LastWriteTimeUtc -gt $RefTimeUtc) { return $true }
        if (-not $p.PSIsContainer) { continue }
        if (Get-ChildItem -LiteralPath $p.FullName -Recurse -ErrorAction SilentlyContinue |
                Where-Object -Property LastWriteTimeUtc -GT $RefTimeUtc |
                Select-Object -First 1) {
            return $true
        }
    }
    return $false
}

# Update archives using 7-Zip
function Update-Archive {
    param ([parameter(Mandatory)][hashtable[]]$Configs)

    foreach ($Config in $Configs) {
        if (-not (Test-TaskCondition $Config)) { continue }
        $Config = Expand-EnvironmentVariables $Config

        [string[]]$srcs = $Config.Sources
        [string]$dst = $Config.Destination

        try {
            # Ensure source and destination are reachable.
            foreach ($s in $srcs + (Split-Path -Parent $dst)) {
                if (-not (Test-Path -LiteralPath $s)) { throw "'${s}' does not exist." }
            }

            # Optional: block if processes are using sources
            if ($Config.CheckProc -and (Test-RunningProcess -Path $srcs)) {
                Write-Log "Archive skipped (sources in use). dst='${dst}'" -Level DEBUG
                continue
            }

            # Only archive if source files are newer than the destination archive
            if ($Config.OnlyIfNewer -and -not (Test-PathModifiedSince -Path $srcs -RefFile $dst)) {
                Write-Log "Archive skipped (no newer sources). dst='${dst}'" -Level DEBUG
                continue
            }

            # Build switch list
            $switches = [System.Collections.Generic.List[string]]@('u', '-up0q0r2x2y2z1w2', '-t7z')
            $switches.AddRange([string[]]$Config.Parameters)
            if ($Config.OnlyIfNewer) { $switches.Add('-stl') }

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

            # 0: OK, 1: warning (non-fatal)
            if ($LASTEXITCODE -gt 1) { throw "7-Zip exited with code ${LASTEXITCODE}." }

            $s = ConvertTo-HumanSize -Bytes (Get-Item -LiteralPath $dst).Length
            Write-Log ("Archive | result=OK code={0} dst='{1}' size={2} sources={3}" -f $LASTEXITCODE, $dst, $s, $srcs.Count) -Level INFO
        }
        catch {
            Write-Log "Archive | result=FAIL dst='${dst}' error=$($_.Exception.Message)" -Level ERROR
        }
    }
}

# Backup general directories
function Backup-Directory {
    param ([parameter(Mandatory)][hashtable[]]$Configs)

    foreach ($Config in $Configs) {
        if (-not (Test-TaskCondition $Config)) { continue }
        $Config = Expand-EnvironmentVariables $Config
        try {
            $src, $dst = Test-SrcAndDestDirs -Source $Config.Source -Destination $Config.Destination
            # Check for running processes
            if ($Config.CheckProc -and (Test-RunningProcess -Path $src)) {
                Write-Log "Directory skipped (source in use). src='${src}' dst='${dst}'" -Level DEBUG
                continue
            }
            robocopy $src $dst $RoboParams $Config.RoboArgs
            if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with exit code ${LASTEXITCODE}. $($RoboCode[$LASTEXITCODE])" }
            Write-Log ("Directory | src='{0}' -> dst='{1}' | code={2} msg='{3}'" -f $src, $dst, $LASTEXITCODE, $RoboCode[$LASTEXITCODE]) -Level INFO
        }
        catch {
            Write-Log ("Directory | result=FAIL src='{0}' dst='{1}' error={2}" -f $Config.Source, $Config.Destination, $_.Exception.Message) -Level ERROR
        }
    }
}

# Backup OneDrive folder
function Backup-OneDrive {
    param ([parameter(Mandatory)][hashtable]$Config)

    if (-not (Test-TaskCondition $Config)) { return }
    $Config = Expand-EnvironmentVariables $Config
    try {
        $src, $dst = Test-SrcAndDestDirs -Source $Config.Source -Destination $Config.Destination
        # Check for running processes
        if ($Config.CheckProc -and (Test-RunningProcess -Path $src)) {
            Write-Log "OneDrive skipped (source in use). src='${src}' dst='${dst}'" -Level DEBUG
            return
        }
        robocopy $src $dst $RoboParams /MT /XA:S
        if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with exit code ${LASTEXITCODE}. $($RoboCode[$LASTEXITCODE])" }
        Write-Log ("OneDrive | src='{0}' -> dst='{1}' | code={2} msg='{3}'" -f $src, $dst, $LASTEXITCODE, $RoboCode[$LASTEXITCODE]) -Level INFO
    }
    catch {
        Write-Log ("OneDrive | result=FAIL src='{0}' dst='{1}' error={2}" -f $Config.Source, $Config.Destination, $_.Exception.Message) -Level ERROR
        return
    }
    # Apply unpinning to the source path.
    if ($Config.AutoUnpin) { Set-UnpinIfNotPinned -Path $src }
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
    param ([parameter(Mandatory)][string]$Path)

    $PINNED = 0x00080000
    $UNPINNED = 0x00100000
    $COMBINED = $UNPINNED -bor $PINNED
    # If neither Unpinned nor Pinned, add Unpinned attribute.
    Get-ChildItem -LiteralPath $Path -Recurse -Attributes !Offline+!ReadOnly | ForEach-Object {
        if (!($_.Attributes -band $COMBINED)) { $_.Attributes = $_.Attributes -bor $UNPINNED }
    }
}

# Backup VMware virtual machines
function Backup-VMware {
    param ([parameter(Mandatory)][hashtable]$Config)

    if (-not (Test-TaskCondition $Config)) { return }
    $Config = Expand-EnvironmentVariables $Config
    try {
        # Check for accessibility
        $src, $dst = Test-SrcAndDestDirs -Source $Config.Source -Destination $Config.Destination
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
                        Write-Log "VMware skip VM (running): '$($dir.Name)'" -Level DEBUG
                        break
                    }
                }
            }
        }
        # Keep VMs that only exist in the destination. Excluding dirs on the right
        # prevents them from being removed.
        if ($Config.KeepExtra) {
            if ($leftDirs.Count -eq $exclusion.Count) {
                Write-Log "VMware skipped (nothing to copy). dst='${dst}'" -Level DEBUG
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
        if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with exit code ${LASTEXITCODE}. $($RoboCode[$LASTEXITCODE])" }
        Write-Log ("VMware | src='{0}' -> dst='{1}' | code={2} msg='{3}'" -f $src, $dst, $LASTEXITCODE, $RoboCode[$LASTEXITCODE]) -Level INFO
    }
    catch {
        Write-Log ("VMware | result=FAIL src='{0}' dst='{1}' error={2}" -f $Config.Source, $Config.Destination, $_.Exception.Message) -Level ERROR
    }
}

# Import Configuration
$Configuration = Read-Configuration

# Execute Backups
Update-Archive $Configuration.Archives
Backup-Directory $Configuration.Directories
Backup-OneDrive $Configuration.OneDrive
Backup-VMware $Configuration.VMware
