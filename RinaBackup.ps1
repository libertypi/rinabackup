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
$RoboParams = @('/MIR', '/J', '/DCOPY:DATX', '/R:3', '/XJ')

function Get-RobocopySummary {
    param([int]$Code)
    if ($Code -eq 0) { return 'No change.' }
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($Code -band 1) { $parts.Add('Copied successfully.') }
    if ($Code -band 2) { $parts.Add('Extras detected.') }
    if ($Code -band 4) { $parts.Add('Mismatches detected.') }
    if ($Code -band 8) { $parts.Add('Some files failed.') }
    if ($Code -band 16) { $parts.Add('Serious error.') }
    $parts -join ' '
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

# Recursively expand environment variables.
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

# Ensure source and destination directories are reachable. Returns them as
# strings.
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

# Tests if any running processes are located within the specified paths. Returns
# the full path of the first matching process, or $null if no match.
function Test-RunningProcess {
    param([Parameter(Mandatory)][string[]]$Path)

    $sep = [IO.Path]::DirectorySeparatorChar
    $procMap = @{}
    $trie = @{}

    foreach ($p in (Get-Process -ErrorAction SilentlyContinue).Path) {
        if ($p) { $procMap[$p.ToLowerInvariant()] = $p }
    }

    foreach ($p in Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue) {
        if (-not $p) { continue }
        $p = $p.Path.ToLowerInvariant()

        if ([System.IO.Directory]::Exists($p)) {
            # Build trie for directories
            $node = $trie
            foreach ($s in $p.Split($sep)) {
                if (-not $s) { continue }
                if (-not $node.ContainsKey($s)) { $node[$s] = @{} }
                $node = $node[$s]
            }
            # Use $sep to mark directory terminal
            $node[$sep] = $null
        }
        elseif ($procMap.ContainsKey($p)) {
            return $procMap[$p]
        }
    }
    if (-not $trie.Count) { return }

    # Check process paths against trie
    foreach ($p in $procMap.Keys) {
        $node = $trie
        foreach ($s in $p.Split($sep)) {
            if (-not $s) { continue }
            if (-not $node.ContainsKey($s)) { break }
            $node = $node[$s]
            if ($node.ContainsKey($sep)) {
                return $procMap[$p]
            }
        }
    }
}

# Checks if any files or directories in specified paths have been modified since
# a reference file's last write time.
function Test-PathModifiedSince {
    param(
        [Parameter(Mandatory)][string[]]$Path,
        [Parameter(Mandatory)][string]$RefFile
    )

    # If the ref file doesn't exist, treat as "out of date".
    if (-not [IO.File]::Exists($RefFile)) { return $true }
    $RefTimeUtc = [IO.File]::GetLastWriteTimeUtc($RefFile)

    $queue = [System.Collections.Generic.Queue[IO.DirectoryInfo]]::new()
    $skipMask = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System -bor [IO.FileAttributes]::ReparsePoint

    foreach ($p in $Path) {
        try {
            if ([IO.File]::GetLastWriteTimeUtc($p) -gt $RefTimeUtc) { return $true }
            if (-not [IO.Directory]::Exists($p)) { continue }
            # BFS traversal
            $queue.Enqueue([IO.DirectoryInfo]::new($p))
            while ($queue.Count) {
                $p = $queue.Dequeue()
                try {
                    foreach ($e in $p.EnumerateFileSystemInfos()) {
                        if ($e.Attributes -band $skipMask) { continue }
                        if ($e.LastWriteTimeUtc -gt $RefTimeUtc) { return $true }
                        if ($e -is [IO.DirectoryInfo]) { $queue.Enqueue($e) }
                    }
                }
                catch {}
            }
        }
        catch {}
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

            # Optional: skip if processes are using sources
            if ($Config.CheckProc -and ($s = Test-RunningProcess -Path $srcs)) {
                Write-Log "Archive skipped (sources in use). proc='${s}' sources=$($srcs.Count) dst='${dst}'" -Level DEBUG
                continue
            }

            # Only archive if source files are newer than the destination archive
            if ($Config.OnlyIfNewer -and -not (Test-PathModifiedSince -Path $srcs -RefFile $dst)) {
                Write-Log "Archive skipped (no newer sources). sources=$($srcs.Count) dst='${dst}'" -Level DEBUG
                continue
            }

            # Build parameter list
            $params = [System.Collections.Generic.List[string]]@('u', '-up0q0r2x2y2z1w2', '-t7z')
            $params.AddRange([string[]]$Config.Parameters)
            if ($Config.OnlyIfNewer) { $params.Add('-stl') }

            # Password
            if (-not [string]::IsNullOrEmpty($Config.Password)) {
                $s = ConvertTo-SecureString -String $Config.Password -ErrorAction Stop
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
                $params.Add('-mhe=on')
                $params.Add("-p$([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))")
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            }

            # Exclusions
            foreach ($s in @($Config.Exclusion)) { if ($s) { $params.Add("-xr!${s}") } }

            # Complete the command line
            $params.Add('--')
            $params.Add($dst)
            $params.AddRange($srcs)

            # Display command line
            $s = foreach ($s in @($Config.SevenZip) + $params) {
                if ($s.StartsWith('-p')) { '-p***' }
                elseif ($s -match "[\s']") { "'{0}'" -f ($s -replace "'", "''") }
                else { $s }
            }
            Write-Log ('Archive | cmd={0}' -f ($s -join ' ')) -Level DEBUG

            # Execute 7-Zip
            & $Config.SevenZip $params

            # 0: OK, 1: warning (non-fatal)
            if ($LASTEXITCODE -gt 1) { throw "7-Zip exited with code ${LASTEXITCODE}." }

            $s = ConvertTo-HumanSize -Bytes (Get-Item -LiteralPath $dst).Length
            Write-Log ("Archive | result=OK code={0} sources={1} dst='{2}' size={3}" -f $LASTEXITCODE, $srcs.Count, $dst, $s) -Level INFO
        }
        catch {
            Write-Log ("Archive | result=FAIL error='{0}' sources={1} dst='{2}'" -f $_.Exception.Message, $srcs.Count, $dst) -Level ERROR
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
            if ($Config.CheckProc -and ($proc = Test-RunningProcess -Path $src)) {
                Write-Log "Directory skipped (source in use). proc='${proc}' src='${src}' dst='${dst}'" -Level DEBUG
                continue
            }

            robocopy $src $dst $RoboParams $Config.RoboArgs
            if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with exit code ${LASTEXITCODE}. ($(Get-RobocopySummary $LASTEXITCODE))" }
            Write-Log ("Directory | result=OK code={0} msg='{1}' | src='{2}' -> dst='{3}'" -f $LASTEXITCODE, (Get-RobocopySummary $LASTEXITCODE), $src, $dst) -Level INFO
        }
        catch {
            Write-Log ("Directory | result=FAIL error='{0}' | src='{1}' -> dst='{2}' " -f $_.Exception.Message, $Config.Source, $Config.Destination) -Level ERROR
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
        if ($Config.CheckProc -and ($proc = Test-RunningProcess -Path $src)) {
            Write-Log "OneDrive skipped (source in use). proc='${proc}' src='${src}' dst='${dst}'" -Level DEBUG
            return
        }

        robocopy $src $dst $RoboParams /MT /XA:S

        if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with exit code ${LASTEXITCODE}. ($(Get-RobocopySummary $LASTEXITCODE))" }
        Write-Log ("OneDrive | result=OK code={0} msg='{1}' | src='{2}' -> dst='{3}'" -f $LASTEXITCODE, (Get-RobocopySummary $LASTEXITCODE), $src, $dst) -Level INFO
    }
    catch {
        Write-Log ("OneDrive | result=FAIL error='{0}' | src='{1}' -> dst='{2}' " -f $_.Exception.Message, $Config.Source, $Config.Destination) -Level ERROR
        return
    }
    # Apply unpinning to the source.
    if ($Config.AutoUnpin) { Set-UnpinIfNotPinned -Path $src }
}

<#
    Frees up space for all items in the OneDrive folder except those manually
    pinned ("Always keep on this device"). If a file is neither PINNED nor
    UNPINNED, it is "locally available." Setting UNPINNED tells the OneDrive
    client to dehydrate it to "online-only."

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
    $pruneMask = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
    $skipMask = $PINNED -bor $UNPINNED -bor [IO.FileAttributes]::Offline

    $stack = [System.Collections.Generic.Stack[IO.DirectoryInfo]]::new()
    $stack.Push([IO.DirectoryInfo]::new($Path))

    while ($stack.Count) {
        $d = $stack.Pop()
        try {
            foreach ($e in $d.EnumerateFileSystemInfos()) {
                $ea = $e.Attributes
                if ($ea -band $pruneMask) { continue }
                if (!($ea -band $skipMask)) { $e.Attributes = $ea -bor $UNPINNED }
                if ($e -is [IO.DirectoryInfo]) { $stack.Push($e) }
            }
        }
        catch {
            Write-Log ("OneDrive | enumerate failed dir='{0}' error='{1}'" -f $d.FullName, $_.Exception.Message) -Level WARNING
        }
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
                        Write-Log "VMware skip VM (running): '$($dir.FullName)'" -Level DEBUG
                        break
                    }
                }
            }
        }
        # Keep VMs that only exist in the destination. Excluding dirs on the
        # right prevents them from being removed.
        if ($Config.KeepExtra) {
            if ($leftDirs.Count -eq $exclusion.Count) {
                Write-Log "VMware skipped (nothing to copy). src='${src}' dst='${dst}'" -Level DEBUG
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
        if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with exit code ${LASTEXITCODE}. ($(Get-RobocopySummary $LASTEXITCODE))" }
        Write-Log ("VMware | result=OK code={0} msg='{1}' | src='{2}' -> dst='{3}'" -f $LASTEXITCODE, (Get-RobocopySummary $LASTEXITCODE), $src, $dst) -Level INFO
    }
    catch {
        Write-Log ("VMware | result=FAIL error='{0}' | src='{1}' -> dst='{2}' " -f $_.Exception.Message, $Config.Source, $Config.Destination) -Level ERROR
    }
}

# Import Configuration
$Configuration = Read-Configuration

# Execute Backups
Update-Archive $Configuration.Archives
Backup-Directory $Configuration.Directories
Backup-OneDrive $Configuration.OneDrive
Backup-VMware $Configuration.VMware
