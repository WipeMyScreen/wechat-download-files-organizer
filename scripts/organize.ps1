#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$RootPath,
    [switch]$ListAccounts,
    [string]$WechatBasePath,
    [switch]$Apply,
    [switch]$Deduplicate,
    [switch]$MigrateLegacyAudio,
    [switch]$AsJson,
    [switch]$IncludeRecords
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Match redirected output in Git Bash, Python, Codex and Claude Code.
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$culture = [Globalization.CultureInfo]::InvariantCulture
$issues = [Collections.Generic.List[object]]::new()
$skippedLinks = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$records = [Collections.Generic.List[object]]::new()
$categories = @('word', 'excel', 'ppt', '压缩包', 'DWG', 'PDF', '音频', '其他')
$extensionGroups = [ordered]@{
    word = '.doc .docx .docm .dot .dotx .dotm .wps .wpt'
    excel = '.xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .et .ett'
    ppt = '.ppt .pptx .pptm .pot .potx .potm .pps .ppsx .ppsm .dps .dpt'
    '压缩包' = '.zip .rar .7z .tar .gz .tgz .bz2 .xz .zst .cab'
    DWG = '.dwg'
    PDF = '.pdf'
    '音频' = '.mp3 .wav .flac .m4a .m4b .aac .ogg .opus .wma .amr .aif .aiff .ape .alac .mid .midi .silk'
}
$extensionMap = @{}
foreach ($group in $extensionGroups.Keys) {
    foreach ($extension in $extensionGroups[$group].Split(' ')) { $extensionMap[$extension] = $group }
}

function Get-NormalizedDirectoryPath([string]$Path) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $pathRoot = [IO.Path]::GetPathRoot($fullPath)
    # Preserve drive and UNC roots; TrimEnd alone would turn C:\ into C:.
    while ($fullPath.Length -gt $pathRoot.Length -and
        ($fullPath.EndsWith('\') -or $fullPath.EndsWith('/'))) {
        $fullPath = $fullPath.Substring(0, $fullPath.Length - 1)
    }
    return $fullPath
}

function Get-BeijingTime {
    [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTimeOffset]::UtcNow, 'China Standard Time').ToString('yyyy-MM-dd HH:mm:ss', $culture)
}

function Assert-PlainDirectory([string]$Path, [switch]$CheckParents) {
    $item = Get-Item -LiteralPath $Path -Force
    if (-not $item.PSIsContainer) { throw "不是文件夹：$Path" }
    while ($null -ne $item) {
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "不能通过目录链接访问：$($item.FullName)" }
        if (-not $CheckParents) { break }
        $item = $item.Parent
    }
}

function Add-Issue([string]$Stage, [string]$Path, [string]$Message) {
    $issues.Add([pscustomobject]@{ Stage = $Stage; Path = $Path; Message = $Message })
}

function Get-RegularFiles([string]$Path, [string]$ExcludedDirectory = '', [string]$Stage = 'Scan') {
    $stack = [Collections.Generic.Stack[string]]::new()
    $stack.Push($Path)
    while ($stack.Count -gt 0) {
        $directory = $stack.Pop()
        try {
            Assert-PlainDirectory $directory
            $children = @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)
        } catch {
            Add-Issue $Stage $directory $_.Exception.Message
            continue
        }
        foreach ($child in $children) {
            if ($ExcludedDirectory -and $child.FullName.Equals($ExcludedDirectory, [StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                [void]$skippedLinks.Add($child.FullName)
                continue
            }
            if ($child.PSIsContainer) { $stack.Push($child.FullName) } else { $child }
        }
    }
}

function Get-FreeDestination([string]$Directory, [string]$Name) {
    $stem = [IO.Path]::GetFileNameWithoutExtension($Name)
    $extension = [IO.Path]::GetExtension($Name)
    $candidate = Join-Path $Directory $Name
    $index = 0
    while ($reserved.Contains($candidate) -or (Test-Path -LiteralPath $candidate)) {
        $index++
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $stem, $index, $extension)
    }
    return $candidate
}

function Format-FileSize([long]$Bytes) {
    $divisor = 1024.0
    $unit = 'KB'
    if ($Bytes -ge 1GB) {
        $divisor = [double](1GB)
        $unit = 'GB'
    } elseif ($Bytes -ge 1MB) {
        $divisor = [double](1MB)
        $unit = 'MB'
    }
    return ($Bytes / $divisor).ToString('F2', $culture) + ' ' + $unit
}

function Get-DisplayWidth([string]$Text) {
    $displayWidth = 0
    $elements = [Globalization.StringInfo]::GetTextElementEnumerator($Text)
    while ($elements.MoveNext()) {
        $element = $elements.GetTextElement()
        $code = [char]::ConvertToUtf32($element, 0)
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($element, 0)
        if ($category -in @([Globalization.UnicodeCategory]::NonSpacingMark, [Globalization.UnicodeCategory]::EnclosingMark, [Globalization.UnicodeCategory]::Format)) { continue }
        if (($code -ge 0x1100 -and $code -le 0x115f) -or
            ($code -ge 0x2e80 -and $code -le 0xa4cf) -or
            ($code -ge 0xac00 -and $code -le 0xd7a3) -or
            ($code -ge 0xf900 -and $code -le 0xfaff) -or
            ($code -ge 0xfe10 -and $code -le 0xfe6f) -or
            ($code -ge 0xff01 -and $code -le 0xff60) -or
            ($code -ge 0xffe0 -and $code -le 0xffe6) -or
            ($code -ge 0x1f1e6 -and $code -le 0x1faff) -or
            ($code -ge 0x20000 -and $code -le 0x3fffd)) {
            $displayWidth += 2
        } else { $displayWidth++ }
    }
    return $displayWidth
}

function Format-Cell([string]$Text, [int]$Width, [ValidateSet('Left', 'Right')][string]$Alignment = 'Left') {
    $padding = ' ' * [Math]::Max(0, $Width - (Get-DisplayWidth $Text))
    if ($Alignment -eq 'Right') { return $padding + $Text }
    return $Text + $padding
}

function Format-ReportTable([object[]]$Rows, [long]$TotalFiles, [string]$TotalSize, [int]$MinimumWidth = 0) {
    $folderWidth = 12
    $countWidth = 8
    $sizeWidth = [Math]::Max(10, (Get-DisplayWidth $TotalSize))
    $countWidth = [Math]::Max($countWidth, (Get-DisplayWidth "$TotalFiles 个"))
    foreach ($row in $Rows) {
        $folderWidth = [Math]::Max($folderWidth, (Get-DisplayWidth $row.Folder))
        $countWidth = [Math]::Max($countWidth, (Get-DisplayWidth "$($row.FileCount) 个"))
        $sizeWidth = [Math]::Max($sizeWidth, (Get-DisplayWidth $row.Size))
    }
    $tableWidth = $folderWidth + $countWidth + $sizeWidth + 6
    if ($tableWidth -lt $MinimumWidth) { $folderWidth += $MinimumWidth - $tableWidth }
    $divider = ('-' * $folderWidth) + '-+-' + ('-' * $countWidth) + '-+-' + ('-' * $sizeWidth)
    (Format-Cell '文件夹' $folderWidth) + ' | ' + (Format-Cell '文件数量' $countWidth -Alignment Right) + ' | ' + (Format-Cell '文件大小' $sizeWidth -Alignment Right)
    $divider
    foreach ($row in $Rows) {
        (Format-Cell $row.Folder $folderWidth) + ' | ' + (Format-Cell "$($row.FileCount) 个" $countWidth -Alignment Right) + ' | ' + (Format-Cell $row.Size $sizeWidth -Alignment Right)
    }
    $divider
    (Format-Cell '总计' $folderWidth) + ' | ' + (Format-Cell "$TotalFiles 个" $countWidth -Alignment Right) + ' | ' + (Format-Cell $TotalSize $sizeWidth -Alignment Right)
}

if ($ListAccounts) {
    if ($Apply -or $Deduplicate -or $MigrateLegacyAudio -or -not [string]::IsNullOrWhiteSpace($RootPath)) {
        throw '列出账号不能与 -Apply、-Deduplicate、-MigrateLegacyAudio 或 -RootPath 同时使用。'
    }
    if ([string]::IsNullOrWhiteSpace($WechatBasePath)) {
        $userProfileDirectory = [Environment]::GetFolderPath('UserProfile')
        $documents = [Environment]::GetFolderPath('MyDocuments')
        $basePaths = @($userProfileDirectory, $documents) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Join-Path $_ 'xwechat_files' }
    } else {
        $basePaths = @($WechatBasePath)
    }
    $searchRoots = [Collections.Generic.List[string]]::new()
    $seenRoots = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in $basePaths) {
        $normalized = Get-NormalizedDirectoryPath $path
        if ($seenRoots.Add($normalized)) { $searchRoots.Add($normalized) }
    }
    $candidates = [Collections.Generic.List[object]]::new()
    foreach ($basePath in $searchRoots) {
        if (-not (Test-Path -LiteralPath $basePath)) { continue }
        Assert-PlainDirectory $basePath -CheckParents
        $accounts = @(Get-ChildItem -LiteralPath $basePath -Directory -Force | Sort-Object Name)
        :accountLoop foreach ($account in $accounts) {
            # Walk only the known account/msg/file shape, never enumerate users' files.
            $candidatePath = $account.FullName
            foreach ($component in @('', 'msg', 'file')) {
                if ($component) { $candidatePath = Join-Path $candidatePath $component }
                try {
                    $candidateItem = Get-Item -LiteralPath $candidatePath -Force -ErrorAction Stop
                } catch [System.Management.Automation.ItemNotFoundException] {
                    continue accountLoop
                } catch {
                    Add-Issue 'Discovery' $candidatePath $_.Exception.Message
                    continue accountLoop
                }
                if ($candidateItem.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                    [void]$skippedLinks.Add($candidatePath)
                    continue accountLoop
                }
                if (-not $candidateItem.PSIsContainer) { continue accountLoop }
            }
            $candidates.Add([pscustomobject]@{
                AccountFolder = $account.Name
                RootPath = $candidatePath
            })
        }
    }
    $discovery = [pscustomobject]@{
        Mode = 'SelectAccount'
        Success = ($issues.Count -eq 0)
        SelectionRequired = ($candidates.Count -ne 1)
        SearchRoots = $searchRoots.ToArray()
        CandidateCount = $candidates.Count
        Candidates = $candidates.ToArray()
        SkippedLinks = @($skippedLinks)
        Failures = $issues.ToArray()
    }
    if ($AsJson) { $discovery | ConvertTo-Json -Depth 6 }
    else {
        '微信账号目录（搜索位置：{0}）' -f ($searchRoots -join '；')
        foreach ($candidate in $candidates) { '{0}：{1}' -f $candidate.AccountFolder, $candidate.RootPath }
        if ($candidates.Count -eq 0) { '未找到有效账号目录，请提供完整的 msg\file 路径，或指定其他 -WechatBasePath。' }
        '仅发现候选目录；唯一候选无需单独选账号，多候选须明确目录。执行时仍需 -RootPath；未移动或修改任何文件。'
        foreach ($issue in $issues) { '[{0}] {1}：{2}' -f $issue.Stage, $issue.Path, $issue.Message }
    }
    return
}
if (-not [string]::IsNullOrWhiteSpace($WechatBasePath)) {
    throw '-WechatBasePath 仅能与 -ListAccounts 一起使用。'
}
if ([string]::IsNullOrWhiteSpace($RootPath)) {
    throw '未选择微信账号目录。请先用 -ListAccounts 列出账号并询问用户，再通过 -RootPath 指定完整的 msg\file 路径。'
}

if ($Deduplicate) { . (Join-Path $PSScriptRoot 'deduplicate.ps1') }
$root = Get-NormalizedDirectoryPath $RootPath
Assert-PlainDirectory $root -CheckParents
$output = Join-Path $root '整理归纳'
$reportPath = Join-Path $output '整理统计.txt'
foreach ($directory in @($output) + @($categories | ForEach-Object { Join-Path $output $_ })) {
    if (Test-Path -LiteralPath $directory) { Assert-PlainDirectory $directory }
}
if (Test-Path -LiteralPath $reportPath) {
    $existingReport = Get-Item -LiteralPath $reportPath -Force
    if ($existingReport.PSIsContainer -or ($existingReport.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw "统计报告位置必须是普通文件：$reportPath"
    }
}

# Serialize runs for this root without creating a lock file during preview.
$hasher = [Security.Cryptography.SHA256]::Create()
try {
    $rootHash = [BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($root.ToUpperInvariant()))).Replace('-', '')
} finally { $hasher.Dispose() }
$mutex = [Threading.Mutex]::new($false, "Local\WechatFileOrganizer_$rootHash")
$acquired = $false
$statistics = [Collections.Generic.List[object]]::new()
$statisticsComplete = $true
$reportWritten = $false
try {
    try { $acquired = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $acquired = $true }
    if (-not $acquired) { throw '此目录已有整理任务运行，请稍后重试。' }

    $incoming = @(Get-RegularFiles -Path $root -ExcludedDirectory $output | Sort-Object FullName)
    $audioMigrations = @()
    $otherDirectory = Join-Path $output '其他'
    if ($MigrateLegacyAudio -and (Test-Path -LiteralPath $otherDirectory)) {
        $audioMigrations = @(Get-RegularFiles -Path $otherDirectory | Where-Object { $extensionMap[$_.Extension.ToLowerInvariant()] -eq '音频' } | Sort-Object FullName)
    }
    $files = @($audioMigrations) + @($incoming)
    $migrationPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $audioMigrations) { [void]$migrationPaths.Add($file.FullName) }
    $newPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $runBytes = @{}
    foreach ($category in $categories) { $runBytes[$category] = [long]0 }
    foreach ($file in $files) {
        $category = $extensionMap[$file.Extension.ToLowerInvariant()]
        if (-not $category) { $category = '其他' }
        $destination = ''
        $record = [pscustomobject]@{
            Time = Get-BeijingTime; Source = $file.FullName; Destination = ''
            Category = $category; Status = 'Planned'; Renamed = $false; Error = ''; KeptPath = $null
        }
        try {
            $targetDirectory = Join-Path $output $category
            $destination = Get-FreeDestination $targetDirectory $file.Name
            if ($Apply) {
                Assert-PlainDirectory $file.DirectoryName -CheckParents
                $current = Get-Item -LiteralPath $file.FullName -Force
                if ($current.PSIsContainer -or ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw '源文件已变成文件夹或链接，已跳过。' }
                [void][IO.Directory]::CreateDirectory($targetDirectory)
                Assert-PlainDirectory $targetDirectory -CheckParents
                [IO.File]::Move($file.FullName, $destination)
                $record.Status = 'Moved'
                if (-not $migrationPaths.Contains($file.FullName)) { [void]$newPaths.Add($destination) }
            }
            [void]$reserved.Add($destination)
            $runBytes[$category] += $(if ($Apply) { $current.Length } else { $file.Length })
        } catch {
            $record.Status = 'Failed'
            $record.Error = $_.Exception.Message
            Add-Issue 'Move' $file.FullName $record.Error
        }
        $record.Destination = $destination
        $record.Renamed = [bool]($destination -and ([IO.Path]::GetFileName($destination) -cne $file.Name))
        $records.Add($record)
    }
    $categoryCounts = [ordered]@{}
    foreach ($category in $categories) {
        $categoryCounts[$category] = @($records | Where-Object { $_.Category -eq $category -and $_.Status -ne 'Failed' }).Count
    }

    $runStatistics = @(foreach ($category in $categories) {
        if ($categoryCounts[$category] -gt 0) {
            [pscustomobject]@{ Folder = $category; FileCount = $categoryCounts[$category]; Size = (Format-FileSize $runBytes[$category]) }
        }
    })
    [long]$runTotalFiles = 0
    [long]$runTotalBytes = 0
    foreach ($category in $categories) {
        $runTotalFiles += $categoryCounts[$category]
        $runTotalBytes += $runBytes[$category]
    }
    $runSummary = [pscustomobject]@{
        Phase = $(if ($Apply) { 'Completed' } else { 'Planned' })
        Statistics = $runStatistics
        TotalFiles = $runTotalFiles
        TotalSize = Format-FileSize $runTotalBytes
        Complete = (@($issues | Where-Object { $_.Stage -in @('Scan', 'Move') }).Count -eq 0)
        AudioMigrationCount = @($records | Where-Object { $_.Status -in @('Moved', 'Planned') -and $migrationPaths.Contains($_.Source) }).Count
    }

    $duplicatePairs = @()
    [long]$duplicateBytes = 0
    [long]$recycledBytes = 0
    $recycledCount = 0
    $duplicateCheckComplete = $null
    if ($Deduplicate) {
        $entries = [Collections.Generic.List[object]]::new()
        foreach ($category in $categories) {
            $directory = Join-Path $output $category
            if (-not (Test-Path -LiteralPath $directory)) { continue }
            foreach ($file in @(Get-RegularFiles -Path $directory -Stage 'Duplicate')) {
                $group = $category
                if (-not $Apply -and $migrationPaths.Contains($file.FullName)) { $group = '音频' }
                $entries.Add((New-DuplicateEntry $file $group (-not $newPaths.Contains($file.FullName))))
            }
        }
        if (-not $Apply) {
            foreach ($file in $incoming) {
                $category = $extensionMap[$file.Extension.ToLowerInvariant()]
                if (-not $category) { $category = '其他' }
                $entries.Add((New-DuplicateEntry $file $category $false))
            }
        }
        $duplicatePairs = @(Get-DuplicatePlan -Entries $entries.ToArray())
        $duplicateCheckComplete = @($issues | Where-Object { $_.Stage -in @('Scan', 'Duplicate', 'Move') }).Count -eq 0
        foreach ($pair in $duplicatePairs) {
            $duplicateBytes += $pair.Candidate.Length
            $detail = [pscustomobject]@{
                Time = Get-BeijingTime; Source = $pair.Candidate.Path; Destination = $null
                Category = $pair.Candidate.Category; Status = 'DuplicatePlanned'
                Renamed = $false; Error = ''; KeptPath = $pair.Keeper.Path
            }
            if ($Apply) {
                try {
                    $null = Invoke-DuplicateRecycle $pair
                    $recycledCount++
                    $recycledBytes += $pair.Candidate.Length
                    $detail.Status = 'Recycled'
                } catch {
                    $detail.Status = 'RecycleFailed'
                    $detail.Error = $_.Exception.Message
                    Add-Issue 'Recycle' $pair.Candidate.Path $detail.Error
                }
            }
            $records.Add($detail)
        }
    }
    if ($Apply) {
        # Non-recursive: a subdirectory, hidden file or a concurrent arrival prevents removal.
        foreach ($category in $categories) {
            $directory = Join-Path $output $category
            if (-not (Test-Path -LiteralPath $directory)) { continue }
            try {
                Assert-PlainDirectory $directory -CheckParents
                if (@(Get-ChildItem -LiteralPath $directory -Force).Count -eq 0) { [IO.Directory]::Delete($directory, $false) }
            } catch { Add-Issue 'EmptyDirectory' $directory $_.Exception.Message }
        }
        $issueCountBeforeStats = $issues.Count
        try {
            if (Test-Path -LiteralPath $output) {
                $folders = @(Get-ChildItem -LiteralPath $output -Directory -Force)
                $orderedNames = @($categories | Where-Object { Test-Path -LiteralPath (Join-Path $output $_) }) + @($folders | Where-Object { $_.Name -notin $categories } | Sort-Object Name | Select-Object -ExpandProperty Name)
                foreach ($name in $orderedNames) {
                    $folder = Get-Item -LiteralPath (Join-Path $output $name) -Force
                    if ($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) { [void]$skippedLinks.Add($folder.FullName); continue }
                    $folderFiles = @(Get-RegularFiles -Path $folder.FullName -Stage 'Statistics')
                    if ($folderFiles.Count -eq 0) { continue }
                    [long]$bytes = 0
                    foreach ($item in $folderFiles) { $bytes += $item.Length }
                    $statistics.Add([pscustomobject]@{ Folder = $name; FileCount = $folderFiles.Count; Size = (Format-FileSize $bytes); Bytes = $bytes })
                }
            }
        } catch { Add-Issue 'Statistics' $output $_.Exception.Message }
        $statisticsComplete = $issues.Count -eq $issueCountBeforeStats
    }
    [long]$totalFiles = 0
    [long]$totalBytes = 0
    foreach ($row in $statistics) { $totalFiles += $row.FileCount; $totalBytes += $row.Bytes }
    $inventory = $null
    if ($Apply) {
        $inventory = [pscustomobject]@{
            Statistics = @($statistics | Select-Object Folder, FileCount, Size)
            TotalFiles = $totalFiles
            TotalSize = Format-FileSize $totalBytes
            Complete = $statisticsComplete
        }
    }
    $recycleFailedCount = @($issues | Where-Object Stage -eq 'Recycle').Count
    if ($Apply) {
        $timestampLine = '统计时间：{0}（北京时间）' -f (Get-BeijingTime)
        $table = @(Format-ReportTable -Rows $statistics.ToArray() -TotalFiles $totalFiles -TotalSize (Format-FileSize $totalBytes) -MinimumWidth (Get-DisplayWidth $timestampLine))
        $outerDivider = '=' * (Get-DisplayWidth $table[0])
        $lines = [Collections.Generic.List[string]]::new()
        $lines.Add('微信文件整理统计')
        $lines.Add($outerDivider)
        $lines.Add($timestampLine)
        $lines.Add(('本次移动成功：{0} 个，大小 {1}；移动失败 {2} 个。' -f $runSummary.TotalFiles, $runSummary.TotalSize, @($records | Where-Object Status -eq 'Failed').Count))
        if (-not $runSummary.Complete) { $lines.Add('注意：本次扫描或移动有失败，请检查失败详情。') }
        if ($MigrateLegacyAudio) { $lines.Add(('其中旧音频迁移成功：{0} 个。' -f $runSummary.AudioMigrationCount)) }
        $lines.Add('以下表格为整理后库存（含历史文件）：')
        if (-not $statisticsComplete) { $lines.Add('注意：部分目录读取失败，以下统计不完整。') }
        $lines.Add('')
        $lines.AddRange([string[]]$table)
        $lines.Add('')
        if ($Deduplicate) {
            $lines.Add(('本次去重：{0} 个重复副本移入回收站，大小 {1}；回收失败 {2} 个。' -f $recycledCount, (Format-FileSize $recycledBytes), $recycleFailedCount))
            if (-not $duplicateCheckComplete) { $lines.Add('注意：部分文件无法完成重复校验，相关文件已保留。') }
        }
        $lines.Add('说明：')
        $lines.Add('统计包含子文件夹中的文件。')
        $lines.Add('本报告不计入统计。')
        $lines.Add('建议使用等宽字体（如新宋体）查看本表。')
        $lines.Add($outerDivider)
        $temporaryReport = Join-Path $output ('.统计_' + [Guid]::NewGuid().ToString('N') + '.tmp')
        try {
            [void][IO.Directory]::CreateDirectory($output)
            Assert-PlainDirectory $output -CheckParents
            [IO.File]::WriteAllLines($temporaryReport, $lines, [Text.UTF8Encoding]::new($true))
            if (Test-Path -LiteralPath $reportPath) { [IO.File]::Replace($temporaryReport, $reportPath, [System.Management.Automation.Language.NullString]::Value) }
            else { [IO.File]::Move($temporaryReport, $reportPath) }
            $reportWritten = $true
        } catch { Add-Issue 'Report' $reportPath $_.Exception.Message }
        finally { if (Test-Path -LiteralPath $temporaryReport) { Remove-Item -LiteralPath $temporaryReport -Force } }
    }
    $result = [pscustomobject]@{
        Mode = $(if ($Apply) { 'Apply' } else { 'Preview' })
        Success = ($issues.Count -eq 0)
        RootPath = $root; OutputPath = $output
        FileCount = $files.Count
        CategoryCounts = $categoryCounts
        RenamedCount = @($records | Where-Object { $_.Renamed -and $_.Status -in @('Moved', 'Planned') }).Count
        FailedCount = @($records | Where-Object Status -eq 'Failed').Count
        SkippedLinks = @($skippedLinks)
        Failures = $issues.ToArray()
        Run = $runSummary
        Inventory = $inventory
        ReportPath = $(if ($reportWritten) { $reportPath } else { $null })
        Deduplicate = [bool]$Deduplicate
        MigrateLegacyAudio = [bool]$MigrateLegacyAudio
        DuplicateCount = $duplicatePairs.Count
        DuplicateSize = Format-FileSize $duplicateBytes
        RecycledCount = $recycledCount
        RecycledSize = Format-FileSize $recycledBytes
        RecycleFailedCount = $recycleFailedCount
        DuplicateCheckComplete = $duplicateCheckComplete
    }
    if ($IncludeRecords) { $result | Add-Member -NotePropertyName Records -NotePropertyValue $records.ToArray() }
    if ($AsJson) { $result | ConvertTo-Json -Depth 8 }
    else {
        if ($Apply) { '整理完成（请检查下列失败信息）。' } else { '只读预览：未移动文件，未写入目录或报告。' }
        if ($Apply) {
            '本次移动成功：{0} 个，大小 {1}；整理后库存：{2} 个，大小 {3}。' -f $runSummary.TotalFiles, $runSummary.TotalSize, $inventory.TotalFiles, $inventory.TotalSize
            if (-not $runSummary.Complete) { '注意：本次扫描或移动有失败，请检查失败详情。' }
            if (-not $inventory.Complete) { '注意：整理后库存统计不完整。' }
        } else {
            '本次预计分类（仅本次待处理文件，不含其余已归类文件）：'
            if (-not $runSummary.Complete) { '注意：部分文件或目录处理失败，以下预览统计不完整。' }
            Format-ReportTable -Rows $runStatistics -TotalFiles $runTotalFiles -TotalSize $runSummary.TotalSize
        }
        if ($MigrateLegacyAudio) { '本次旧音频迁移（计划或成功）：{0} 个' -f $runSummary.AudioMigrationCount }
        '重命名：{0} 个；移动失败：{1} 个；跳过链接：{2} 个' -f $result.RenamedCount, $result.FailedCount, $skippedLinks.Count
        if ($IncludeRecords) { foreach ($record in $records) { '[{0}] {1} -> {2}' -f $record.Status, $record.Source, $record.Destination; if ($record.KeptPath) { '保留：{0}' -f $record.KeptPath } } }
        if ($Deduplicate) {
            '发现重复副本：{0} 个，大小 {1}；已移入回收站：{2} 个，大小 {3}；回收失败：{4} 个。' -f $duplicatePairs.Count, $result.DuplicateSize, $recycledCount, $result.RecycledSize, $recycleFailedCount
            if (-not $duplicateCheckComplete) { '注意：重复校验不完整，无法确认的文件已保留。' }
        }
        foreach ($link in $skippedLinks) { '跳过链接：{0}' -f $link }
        foreach ($issue in $issues) { '[{0}] {1}：{2}' -f $issue.Stage, $issue.Path, $issue.Message }
        if ($reportWritten) { "统计报告：$reportPath" }
    }
} finally {
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
