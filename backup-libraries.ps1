<#
.SYNOPSIS
  Uses Robocopy to back up the physical folders referenced by Windows Libraries.

.DESCRIPTION
  This script discovers local filesystem folders referenced by .library-ms files in the
  current user's Windows Libraries folder, then backs those folders up to a chosen backup
  root using robocopy.

  I made this happen when Windows File History turned very flaky. It doesn't work in quite
  the same way, but brings a similar outcome after running.

  Use -Help for help.

  The script supports three modes:
    - All:               copy new/changed files, then process files deleted from the source.
    - CopyOnly:          copy new/changed files only.
    - DeletedSweepOnly:  process files deleted from the source only.

  Deleted-source files are not immediately destroyed. If a file exists in the backup but
  no longer exists in the source, it is moved within the backup root to _deleted and later
  pruned after -DeletedRetentionDays days.

  Safety-focused behaviour:
    - Parameters are non-positional: you must use -BackupRoot explicitly.
    - -Help and --help print usage instead of being treated as a destination.
    - Refuses to use a drive root such as A:\ as the backup root; use A:\Backups\Libraries.
    - Refuses to run if the backup root is inside a source folder, or a source folder is inside
      the backup root, preventing self-copy/recursive backup situations.
    - Refuses same-volume backups by default; use -AllowSameVolume only if intentional.
    - Aborts if any discovered source folder is missing, unless -SkipMissingSources is supplied.
    - Does not use robocopy /MIR, so automatic deletion of missing files is not immediate.
    - Dry-run mode uses robocopy /L and reports moves/deletions without performing them.

.NOTES
  Robocopy exit codes 0-7 are treated as success; 8 or above is treated as failure.
  The script copies file data, attributes, and timestamps using /COPY:DAT and directory
  metadata using /DCOPY:DAT. It does not copy NTFS ACLs or owners or auditing.

.EXAMPLE
  Show help:
    .\backup-libraries.ps1 -Help
    .\backup-libraries.ps1 --help

.EXAMPLE
  Dry-run everything:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -DryRun

.EXAMPLE
  Initial or housekeeping run: copy new/changed files and process deleted-source files:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -Mode All

.EXAMPLE
  Routine lightweight run: copy only new/changed files:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -Mode CopyOnly

.EXAMPLE
  Deleted-file processing only:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -Mode DeletedSweepOnly
#>

[CmdletBinding(PositionalBinding = $false)]
param(
    [string] $BackupRoot,

    [ValidateSet("All", "CopyOnly", "DeletedSweepOnly")]
    [string] $Mode = "All",

    [ValidateRange(0, 36500)]
    [int] $DeletedRetentionDays = 31,

    [string] $LibraryDirectory = (Join-Path $env:APPDATA "Microsoft\Windows\Libraries"),

    [string[]] $AdditionalSourceRoots = @(),

    [switch] $SkipKnownFolders,

    [string] $KnownFolderProfileRoot = "",

    [ValidateRange(1, 128)]
    [int] $RobocopyThreads = 1,

    [switch] $NoUnbufferedIo,

    [switch] $NoEta,

    [switch] $AllowSameVolume,

    [switch] $SkipMissingSources,

    [switch] $DryRun,

    [Alias("h", "?")]
    [switch] $Help,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $RemainingArguments
)

$ErrorActionPreference = "Stop"

function Show-Usage {
    $Usage = @"
Library-based robocopy backup script

USAGE
  .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" [options]

COMMON COMMANDS
  Help:
    .\backup-libraries.ps1 -Help
    .\backup-libraries.ps1 --help

  Dry run:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -DryRun

  Copy new/changed files only:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -Mode CopyOnly

  Copy new/changed files and process deleted-source files:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -Mode All

  Deleted-file processing only:
    .\backup-libraries.ps1 -BackupRoot "A:\Backups\Libraries" -Mode DeletedSweepOnly

IMPORTANT SAFETY NOTES
  - -BackupRoot is deliberately not positional. You must name it explicitly.
  - Do not set -BackupRoot to a drive root such as A:\. Use a subfolder.
  - The script refuses to copy onto itself or into a source tree.
  - The script refuses same-volume backups by default; use -AllowSameVolume only if intentional.
  - The script does not use robocopy /MIR.
  - Deleted-source files are moved only inside the backup root, under _deleted.

OPTIONS
  -BackupRoot <path>              Required destination root, e.g. A:\Backups\Libraries
  -Mode <mode>                    All, CopyOnly, or DeletedSweepOnly. Default: All
  -DeletedRetentionDays <days>    Default: 31. Use 0 to disable pruning of _deleted files.
  -LibraryDirectory <path>        Where .library-ms files are read from
  -AdditionalSourceRoots <paths>  Extra source folders to include
  -SkipKnownFolders               Ignore knownfolder:{GUID} entries
  -KnownFolderProfileRoot <path>  Force common known folders under a profile root, e.g. E:\Users\john
  -RobocopyThreads <n>            Default: 1. For large HDD media files, 1 or 2 is often best.
  -NoUnbufferedIo                 Omit robocopy /J
  -NoEta                          Omit robocopy /ETA
  -AllowSameVolume                Permit backup root on same volume as a source
  -SkipMissingSources             Skip missing source folders instead of aborting
  -DryRun                         List/report only; no copying/moving/deleting
  -Help                           Show this help
"@
    Write-Host $Usage
}

if ($Help -or ($RemainingArguments -contains "--help") -or ($RemainingArguments -contains "/?")) {
    Show-Usage
    exit 0
}

if ($RemainingArguments -and $RemainingArguments.Count -gt 0) {
    throw "Unexpected argument(s): $($RemainingArguments -join ' '). Use -Help for usage. Parameters are non-positional; use -BackupRoot explicitly."
}

if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
    Show-Usage
    throw "Missing required parameter: -BackupRoot. Parameters are non-positional, so use -BackupRoot explicitly."
}

if ($BackupRoot.TrimStart().StartsWith("-")) {
    throw "BackupRoot '$BackupRoot' looks like an option, not a path. Use -Help for usage."
}

$KnownFolderRelativeMap = @{
    "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}" = "Documents"
    "{374DE290-123F-4565-9164-39C4925E467B}" = "Downloads"
    "{4BD8D571-6D19-48D3-BE97-422220080E43}" = "Music"
    "{33E28130-4E1E-4676-835A-98395C3BC3BB}" = "Pictures"
    "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}" = "Videos"
    "{AB5FB87B-7CE2-4F83-915D-550846C9537B}" = "Pictures\Camera Roll"
    "{3B193882-D3AD-4EAB-965A-69829D1FB59F}" = "Pictures\Saved Pictures"
}

function Write-Section {
    param([Parameter(Mandatory)] [string] $Text)
    Write-Host ""
    Write-Host "=== $Text ===" -ForegroundColor Cyan
}

function Get-FullPathNormalized {
    param([Parameter(Mandatory)] [string] $Path)
    return [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
}

function Test-IsUnderRoot {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Root
    )

    $p = Get-FullPathNormalized $Path
    $r = Get-FullPathNormalized $Root

    if ($p.Equals($r, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }

    $sep = [System.IO.Path]::DirectorySeparatorChar
    $rWithSlash = $r + $sep
    return $p.StartsWith($rWithSlash, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-RelativePathSafe {
    param(
        [Parameter(Mandatory)] [string] $BasePath,
        [Parameter(Mandatory)] [string] $FullPath
    )

    $Base = [System.IO.Path]::GetFullPath($BasePath)
    $Full = [System.IO.Path]::GetFullPath($FullPath)

    if (-not $Base.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $Base += [System.IO.Path]::DirectorySeparatorChar
    }

    $BaseUri = [System.Uri]::new($Base)
    $FullUri = [System.Uri]::new($Full)
    $RelativeUri = $BaseUri.MakeRelativeUri($FullUri)
    return [System.Uri]::UnescapeDataString($RelativeUri.ToString()).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
}

function Get-ShortHash {
    param([Parameter(Mandatory)] [string] $Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant())
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash).Replace("-", "")).Substring(0, 8)
    }
    finally {
        $sha.Dispose()
    }
}

function Get-SafeJobNameBase {
    param([Parameter(Mandatory)] [string] $SourcePath)

    $full = Get-FullPathNormalized $SourcePath
    $root = [System.IO.Path]::GetPathRoot($full)
    $name = $full

    if ($root) {
        $drive = $root.TrimEnd('\').TrimEnd(':')
        $rest = $full.Substring($root.Length)
        $name = if ($rest) { "$drive`_$rest" } else { $drive }
    }

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + [char[]]@(':', '\', '/', ' ')
    foreach ($ch in $invalidChars) {
        $name = $name.Replace([string]$ch, "_")
    }

    $name = $name -replace '_+', '_'
    $name = $name.Trim('_')
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Source" }
    return $name
}

function Get-KnownFolderPathFromShell {
    param([Parameter(Mandatory)] [string] $GuidText)

    if (-not ("KnownFolderApi" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class KnownFolderApi
{
    [DllImport("shell32.dll")]
    public static extern int SHGetKnownFolderPath(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
        uint dwFlags,
        IntPtr hToken,
        out IntPtr ppszPath);

    [DllImport("ole32.dll")]
    public static extern void CoTaskMemFree(IntPtr pv);
}
"@
    }

    $ptr = [IntPtr]::Zero
    try {
        $guid = [Guid]::Parse($GuidText.Trim('{}'))
        $hr = [KnownFolderApi]::SHGetKnownFolderPath($guid, 0, [IntPtr]::Zero, [ref]$ptr)
        if ($hr -ne 0 -or $ptr -eq [IntPtr]::Zero) { return $null }
        return [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
    }
    catch {
        return $null
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) { [KnownFolderApi]::CoTaskMemFree($ptr) }
    }
}

function Resolve-LibraryLocation {
    param([Parameter(Mandatory)] [string] $Location)

    $loc = $Location.Trim()
    if ([string]::IsNullOrWhiteSpace($loc)) { return $null }

    if ($loc -match '^knownfolder:(\{[0-9A-Fa-f\-]+\})$') {
        if ($SkipKnownFolders) { return $null }

        $guid = $Matches[1].ToUpperInvariant()

        if (-not [string]::IsNullOrWhiteSpace($KnownFolderProfileRoot) -and $KnownFolderRelativeMap.ContainsKey($guid)) {
            return Join-Path $KnownFolderProfileRoot $KnownFolderRelativeMap[$guid]
        }

        $resolved = Get-KnownFolderPathFromShell -GuidText $guid
        if ($resolved) { return $resolved }

        return $null
    }

    if ($loc -match '^file:///(.+)$') {
        $pathPart = [System.Uri]::UnescapeDataString($Matches[1]).Replace('/', '\')
        return $pathPart
    }

    # Local drive path such as E:\Folder. In regex, a literal backslash must be escaped as \\.
    if ($loc -match '^[A-Za-z]:\\') { return $loc }

    # UNC path such as \\server\share\folder.
    if ($loc -match '^\\\\') { return $loc }

    return $null
}

function Get-LibrarySourceRoots {
    param([Parameter(Mandatory)] [string] $Directory)

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        throw "Library directory not found: $Directory"
    }

    $locations = New-Object System.Collections.Generic.List[string]
    $libraryFiles = Get-ChildItem -LiteralPath $Directory -Filter "*.library-ms" -File -ErrorAction Stop

    foreach ($libraryFile in $libraryFiles) {
        Write-Host "Reading library: $($libraryFile.Name)"
        try {
            [xml]$xml = Get-Content -LiteralPath $libraryFile.FullName -Raw
            $nodes = $xml.SelectNodes("//*[local-name()='url']")
            foreach ($node in $nodes) {
                $raw = [string]$node.InnerText
                $resolved = Resolve-LibraryLocation -Location $raw
                if ($resolved) {
                    $locations.Add($resolved)
                }
                else {
                    Write-Warning "Unresolved library location '$raw'. If this is an important source, add it via -AdditionalSourceRoots."
                }
            }
        }
        catch {
            Write-Warning "Could not read library '$($libraryFile.FullName)': $($_.Exception.Message)"
        }
    }

    foreach ($extra in $AdditionalSourceRoots) {
        if (-not [string]::IsNullOrWhiteSpace($extra)) { $locations.Add($extra) }
    }

    $normalized = @()
    foreach ($loc in $locations) {
        try {
            $normalized += (Get-FullPathNormalized $loc)
        }
        catch {
            Write-Warning "Could not normalize source path '$loc': $($_.Exception.Message)"
        }
    }

    return $normalized | Sort-Object -Unique
}

function Assert-BackupRootSafe {
    param(
        [Parameter(Mandatory)] [string] $Root,
        [Parameter(Mandatory)] [string[]] $Sources
    )

    $fullRoot = Get-FullPathNormalized $Root
    $driveRoot = [System.IO.Path]::GetPathRoot($fullRoot).TrimEnd('\')
    $rootTrimmed = $fullRoot.TrimEnd('\')

    if ($rootTrimmed.Equals($driveRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to use a drive root as BackupRoot: '$Root'. Use a subfolder such as A:\Backups\Libraries."
    }

    foreach ($source in $Sources) {
        $sourceFull = Get-FullPathNormalized $source

        if (Test-IsUnderRoot -Path $fullRoot -Root $sourceFull) {
            throw "Unsafe configuration: BackupRoot '$fullRoot' is inside source folder '$sourceFull'. This would copy files onto themselves or recurse."
        }

        if (Test-IsUnderRoot -Path $sourceFull -Root $fullRoot) {
            throw "Unsafe configuration: source folder '$sourceFull' is inside BackupRoot '$fullRoot'. This would back up the backup."
        }

        $sourceDrive = [System.IO.Path]::GetPathRoot($sourceFull)
        $backupDrive = [System.IO.Path]::GetPathRoot($fullRoot)
        if (-not $AllowSameVolume -and $sourceDrive.Equals($backupDrive, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing same-volume backup: source '$sourceFull' and BackupRoot '$fullRoot' are both on '$sourceDrive'. Use -AllowSameVolume only if this is intentional."
        }
    }
}

function New-BackupJobs {
    param(
        [Parameter(Mandatory)] [string[]] $Sources,
        [Parameter(Mandatory)] [string] $Root
    )

    $jobs = New-Object System.Collections.Generic.List[object]
    $usedNames = @{}

    foreach ($source in $Sources) {
        $base = Get-SafeJobNameBase -SourcePath $source
        $name = $base
        if ($usedNames.ContainsKey($name)) {
            $name = "$base`_$(Get-ShortHash $source)"
        }
        $usedNames[$name] = $true

        $dest = Join-Path $Root $name
        if (-not (Test-IsUnderRoot -Path $dest -Root $Root)) {
            throw "Internal safety failure: destination '$dest' is not under BackupRoot '$Root'."
        }

        if (Test-IsUnderRoot -Path $dest -Root $source) {
            throw "Unsafe job: destination '$dest' is inside source '$source'."
        }

        $jobs.Add([pscustomobject]@{
            Name = $name
            Source = $source
            Destination = $dest
        })
    }

    return $jobs
}

function Move-DeletedDestinationFiles {
    param(
        [Parameter(Mandatory)] [string] $JobName,
        [Parameter(Mandatory)] [string] $SourceRoot,
        [Parameter(Mandatory)] [string] $DestinationRoot,
        [Parameter(Mandatory)] [string] $DeletedRoot
    )

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
        throw "Refusing deleted-file sweep because source is missing: $SourceRoot"
    }

    if (-not (Test-Path -LiteralPath $DestinationRoot -PathType Container)) {
        Write-Host "  Destination does not exist yet; nothing to sweep."
        return
    }

    Write-Host "  Deleted-file sweep: checking destination files" -ForegroundColor DarkCyan

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $destFiles = Get-ChildItem -LiteralPath $DestinationRoot -Recurse -File -Force -ErrorAction SilentlyContinue

    foreach ($destFile in $destFiles) {
        $relative = Get-RelativePathSafe -BasePath $DestinationRoot -FullPath $destFile.FullName
        $sourceFile = Join-Path $SourceRoot $relative

        if (-not (Test-Path -LiteralPath $sourceFile -PathType Leaf)) {
            $archivePath = Join-Path $DeletedRoot (Join-Path $JobName (Join-Path $timestamp $relative))
            $archiveDir = Split-Path -Parent $archivePath

            if ($DryRun) {
                Write-Host "DRY RUN: would move deleted-source backup file '$($destFile.FullName)' to '$archivePath'" -ForegroundColor Yellow
            }
            else {
                New-Item -ItemType Directory -Force -Path $archiveDir | Out-Null
                Move-Item -LiteralPath $destFile.FullName -Destination $archivePath -Force
            }
        }
    }
}

function Invoke-DeletedPrune {
    param([Parameter(Mandatory)] [string] $DeletedRoot)

    if ($DeletedRetentionDays -le 0) { return }
    if (-not (Test-Path -LiteralPath $DeletedRoot -PathType Container)) { return }

    $cutoff = (Get-Date).AddDays(-$DeletedRetentionDays)
    Write-Host "Pruning _deleted files older than $DeletedRetentionDays days" -ForegroundColor DarkCyan

    if ($DryRun) {
        Get-ChildItem -LiteralPath $DeletedRoot -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            ForEach-Object { Write-Host "DRY RUN: would delete old _deleted file '$($_.FullName)'" -ForegroundColor Yellow }
        return
    }

    Get-ChildItem -LiteralPath $DeletedRoot -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    Get-ChildItem -LiteralPath $DeletedRoot -Directory -Recurse -Force -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending |
        Where-Object { -not (Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Invoke-RobocopyJob {
    param(
        [Parameter(Mandatory)] [string] $JobName,
        [Parameter(Mandatory)] [string] $Source,
        [Parameter(Mandatory)] [string] $Destination,
        [Parameter(Mandatory)] [string] $LogRoot
    )

    if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
        throw "Source folder not found: $Source"
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logPath = Join-Path $LogRoot "$($timestamp)_$JobName.log"

    $options = @(
        "/E",
        "/XJ",
        "/COPY:DAT",
        "/DCOPY:DAT",
        "/R:2",
        "/W:5",
        "/MT:$RobocopyThreads",
        "/FFT",
        "/TEE",
        "/LOG+:$logPath"
    )

    if (-not $NoUnbufferedIo) { $options += "/J" }
    if (-not $NoEta) { $options += "/ETA" }
    if ($DryRun) { $options += "/L" }

    Write-Host "robocopy '$Source' -> '$Destination'"
    if ($DryRun) { Write-Host "DRY RUN: robocopy will list actions only; no files will be copied." -ForegroundColor Yellow }

    & robocopy $Source $Destination @options
    $exitCode = $LASTEXITCODE

    if ($exitCode -ge 8) {
        throw "robocopy failed for '$Source' with exit code $exitCode. See log: $logPath"
    }
    else {
        Write-Host "robocopy completed for '$Source' with exit code $exitCode. See log: $logPath" -ForegroundColor Green
    }
}

$BackupRoot = Get-FullPathNormalized $BackupRoot
$DeletedRoot = Join-Path $BackupRoot "_deleted"
$LogRoot = Join-Path $BackupRoot "_logs"
$StateRoot = Join-Path $BackupRoot "_state"

Write-Section "Discovering library source folders"
$allSources = @(Get-LibrarySourceRoots -Directory $LibraryDirectory)

$missingSources = @()
$existingSources = @()
foreach ($source in $allSources) {
    if (Test-Path -LiteralPath $source -PathType Container) {
        $existingSources += (Get-FullPathNormalized $source)
    }
    else {
        $missingSources += $source
    }
}

if ($missingSources.Count -gt 0) {
    Write-Warning "Missing source folder(s):"
    $missingSources | ForEach-Object { Write-Warning "  $_" }
    if (-not $SkipMissingSources) {
        throw "Aborting because one or more source folders are missing. This prevents accidental deleted-file sweeps after a drive/path problem. Use -SkipMissingSources only if intentional."
    }
}

if ($existingSources.Count -eq 0) {
    throw "No existing source folders found. Aborting."
}

Assert-BackupRootSafe -Root $BackupRoot -Sources $existingSources

$jobs = @(New-BackupJobs -Sources $existingSources -Root $BackupRoot)

Write-Host "Source folders to back up:"
foreach ($job in $jobs) {
    Write-Host "  $($job.Source) -> $($job.Destination)"
}

if ($DryRun) {
    Write-Host ""
    Write-Host "DRY RUN ENABLED: no files will be copied, moved, or deleted." -ForegroundColor Yellow
}

# Create administrative directories after safety validation. Robocopy cannot create the log parent folder itself.
New-Item -ItemType Directory -Force -Path $BackupRoot, $LogRoot, $StateRoot | Out-Null
if ($Mode -in @("All", "DeletedSweepOnly")) {
    New-Item -ItemType Directory -Force -Path $DeletedRoot | Out-Null
}

$markerPath = Join-Path $StateRoot "backup-libraries.marker.json"
if (-not $DryRun -and -not (Test-Path -LiteralPath $markerPath)) {
    [pscustomobject]@{
        CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
        ComputerName = $env:COMPUTERNAME
        BackupRoot = $BackupRoot
        Script = "backup-libraries.ps1"
    } | ConvertTo-Json | Set-Content -LiteralPath $markerPath -Encoding UTF8
}

foreach ($job in $jobs) {
    Write-Section "Job: $($job.Name)"
    Write-Host "Source:      $($job.Source)"
    Write-Host "Destination: $($job.Destination)"

    if ($Mode -in @("All", "DeletedSweepOnly")) {
        Move-DeletedDestinationFiles -JobName $job.Name -SourceRoot $job.Source -DestinationRoot $job.Destination -DeletedRoot $DeletedRoot
    }

    if ($Mode -in @("All", "CopyOnly")) {
        Invoke-RobocopyJob -JobName $job.Name -Source $job.Source -Destination $job.Destination -LogRoot $LogRoot
    }
}

if ($Mode -in @("All", "DeletedSweepOnly")) {
    Write-Section "Pruning deleted-file archive"
    Invoke-DeletedPrune -DeletedRoot $DeletedRoot
}

Write-Section "Done"
Write-Host "Backup root: $BackupRoot"
Write-Host "Mode:        $Mode"
if ($DryRun) { Write-Host "Dry run:     yes" } else { Write-Host "Dry run:     no" }
