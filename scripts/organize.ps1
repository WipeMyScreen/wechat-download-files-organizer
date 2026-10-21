#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$RootPath,
    [switch]$ListAccounts,
    [string]$WechatBasePath,
    [switch]$Apply,
    [switch]$AsJson
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
$categories = @('word', 'excel', 'ppt', '压缩包', 'DWG', 'PDF', '其他')
$extensionGroups = [ordered]@{
    word = '.doc .docx .docm .dot .dotx .dotm .wps .wpt'
    excel = '.xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .et .ett'
    ppt = '.ppt .pptx .pptm .pot .potx .potm .pps .ppsx .ppsm .dps .dpt'
    '压缩包' = '.zip .rar .7z .tar .gz .tgz .bz2 .xz .zst .cab'
    DWG = '.dwg'
    PDF = '.pdf'
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

function Assert-PlainDirectory([string]$Path) {
    $item = Get-Item -LiteralPath $Path -Force
    if (-not $item.PSIsContainer) { throw "不是文件夹：$Path" }
    while ($null -ne $item) {
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "不能通过目录链接访问：$($item.FullName)" }
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
    if ($Apply -or -not [string]::IsNullOrWhiteSpace($RootPath)) {
        throw '列出账号不能与 -Apply 或 -RootPath 同时使用。请先选择账号，再执行整理。'
    }
    if ([string]::IsNullOrWhiteSpace($WechatBasePath)) {
        $documents = [Environment]::GetFolderPath('MyDocuments')
        if ([string]::IsNullOrWhiteSpace($documents)) { throw '无法定位文档目录，请通过 -WechatBasePath 指定 xwechat_files 路径。' }
        $WechatBasePath = Join-Path $documents 'xwechat_files'
    }
    $basePath = Get-NormalizedDirectoryPath $WechatBasePath
    $candidates = [Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $basePath) {
        Assert-PlainDirectory $basePath
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
        SelectionRequired = $true
        SearchRoot = $basePath
        CandidateCount = $candidates.Count
        Candidates = $candidates.ToArray()
        SkippedLinks = @($skippedLinks)
        Failures = $issues.ToArray()
    }
    if ($AsJson) { $discovery | ConvertTo-Json -Depth 6 }
    else {
        "微信账号目录（搜索位置：$basePath）"
        foreach ($candidate in $candidates) { '{0}：{1}' -f $candidate.AccountFolder, $candidate.RootPath }
        if ($candidates.Count -eq 0) { '未找到有效账号目录，请提供完整的 msg\file 路径，或指定其他 -WechatBasePath。' }
        '请先由用户选择目录，再以 -RootPath 指定完整的 msg\file 路径；未移动或修改任何文件。'
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

$root = Get-NormalizedDirectoryPath $RootPath
Assert-PlainDirectory $root
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
    $files = @(Get-RegularFiles -Path $root -ExcludedDirectory $output | Sort-Object FullName)
    $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($Apply) {
        foreach ($directory in @($output) + @($categories | ForEach-Object { Join-Path $output $_ })) {
            [void][IO.Directory]::CreateDirectory($directory)
            Assert-PlainDirectory $directory
        }
    }

    foreach ($file in $files) {
        $category = $extensionMap[$file.Extension.ToLowerInvariant()]
        if (-not $category) { $category = '其他' }
        $destination = ''
        $record = [pscustomobject]@{
            Time = Get-BeijingTime; Source = $file.FullName; Destination = ''
            Category = $category; Status = 'Planned'; Renamed = $false; Error = ''
        }
        try {
            $targetDirectory = Join-Path $output $category
            $destination = Get-FreeDestination $targetDirectory $file.Name
            if ($Apply) {
                Assert-PlainDirectory $file.DirectoryName
                Assert-PlainDirectory $targetDirectory
                $current = Get-Item -LiteralPath $file.FullName -Force
                if ($current.PSIsContainer -or ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                    throw '源文件已变成文件夹或链接，已跳过。'
                }
                while ($true) {
                    try {
                        # The two-argument overload never overwrites an existing file.
                        [IO.File]::Move($file.FullName, $destination)
                        break
                    } catch [IO.IOException] {
                        if ((Test-Path -LiteralPath $destination) -and (Test-Path -LiteralPath $file.FullName)) {
                            $destination = Get-FreeDestination $targetDirectory $file.Name
                        } else { throw }
                    }
                }
                $record.Status = 'Moved'
            }
            [void]$reserved.Add($destination)
        } catch {
            $record.Status = 'Failed'
            $record.Error = $_.Exception.Message
            Add-Issue 'Move' $file.FullName $record.Error
        }
        $record.Destination = $destination
        $record.Renamed = [bool]($destination -and ([IO.Path]::GetFileName($destination) -cne $file.Name))
        $records.Add($record)
    }

    if ($Apply) {
        $issueCountBeforeStats = $issues.Count
        try {
            $folders = @(Get-ChildItem -LiteralPath $output -Directory -Force)
            $orderedNames = @($categories) + @($folders.Name | Where-Object { $_ -notin $categories } | Sort-Object)
            foreach ($name in $orderedNames) {
                $folder = Get-Item -LiteralPath (Join-Path $output $name) -Force
                if ($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                    [void]$skippedLinks.Add($folder.FullName)
                    continue
                }
                $folderFiles = @(Get-RegularFiles -Path $folder.FullName -Stage 'Statistics')
                [long]$bytes = 0
                foreach ($item in $folderFiles) { $bytes += $item.Length }
                $statistics.Add([pscustomobject]@{ Folder = $name; FileCount = $folderFiles.Count; SizeKB = ($bytes / 1024.0).ToString('F2', $culture); Size = (Format-FileSize $bytes); Bytes = $bytes })
            }
        } catch { Add-Issue 'Statistics' $output $_.Exception.Message }
        $statisticsComplete = $issues.Count -eq $issueCountBeforeStats
        [long]$totalFiles = 0
        [long]$totalBytes = 0
        foreach ($row in $statistics) {
            $totalFiles += $row.FileCount
            $totalBytes += $row.Bytes
        }
        $timestampLine = '统计时间：{0}（北京时间）' -f (Get-BeijingTime)
        $table = @(Format-ReportTable -Rows $statistics.ToArray() -TotalFiles $totalFiles -TotalSize (Format-FileSize $totalBytes) -MinimumWidth (Get-DisplayWidth $timestampLine))
        $outerDivider = '=' * (Get-DisplayWidth $table[0])
        $lines = [Collections.Generic.List[string]]::new()
        $lines.Add('微信文件整理统计')
        $lines.Add($outerDivider)
        $lines.Add($timestampLine)
        if (-not $statisticsComplete) { $lines.Add('注意：部分目录读取失败，以下统计不完整。') }
        $lines.Add('')
        $lines.AddRange([string[]]$table)
        $lines.Add('')
        $lines.Add('说明：')
        $lines.Add('统计包含子文件夹中的文件。')
        $lines.Add('本报告不计入统计。')
        $lines.Add('建议使用等宽字体（如新宋体）查看本表。')
        $lines.Add($outerDivider)
        $temporaryReport = Join-Path $output ('.统计_' + [Guid]::NewGuid().ToString('N') + '.tmp')
        try {
            Assert-PlainDirectory $output
            [IO.File]::WriteAllLines($temporaryReport, $lines, [Text.UTF8Encoding]::new($true))
            if (Test-Path -LiteralPath $reportPath) {
                [IO.File]::Replace($temporaryReport, $reportPath, [System.Management.Automation.Language.NullString]::Value)
            } else {
                [IO.File]::Move($temporaryReport, $reportPath)
            }
            $reportWritten = $true
        } catch { Add-Issue 'Report' $reportPath $_.Exception.Message }
        finally { if (Test-Path -LiteralPath $temporaryReport) { Remove-Item -LiteralPath $temporaryReport -Force } }
    }

    $categoryCounts = [ordered]@{}
    foreach ($category in $categories) {
        $categoryCounts[$category] = @($records | Where-Object { $_.Category -eq $category -and $_.Status -ne 'Failed' }).Count
    }
    $result = [pscustomobject]@{
        Mode = $(if ($Apply) { 'Apply' } else { 'Preview' })
        Success = ($issues.Count -eq 0)
        RootPath = $root; OutputPath = $output
        FileCount = $files.Count
        CategoryCounts = $categoryCounts
        RenamedCount = @($records | Where-Object { $_.Renamed -and $_.Status -ne 'Failed' }).Count
        FailedCount = @($records | Where-Object Status -eq 'Failed').Count
        SkippedLinks = @($skippedLinks)
        Failures = $issues.ToArray()
        Records = $records.ToArray()
        Statistics = @($statistics | Select-Object Folder, FileCount, SizeKB, Size)
        StatisticsComplete = $(if ($Apply) { $statisticsComplete } else { $null })
        ReportPath = $(if ($reportWritten) { $reportPath } else { $null })
    }
    if ($AsJson) { $result | ConvertTo-Json -Depth 8 }
    else {
        if ($Apply) { '整理完成（请检查下列失败信息）。' } else { '只读预览：未移动文件，未写入目录或报告。' }
        foreach ($category in $categories) { '{0}：{1} 个' -f $category, $categoryCounts[$category] }
        '重命名：{0} 个；移动失败：{1} 个；跳过链接：{2} 个' -f $result.RenamedCount, $result.FailedCount, $skippedLinks.Count
        if (-not $Apply) { foreach ($record in $records) { '{0} -> {1}' -f $record.Source, $record.Destination } }
        foreach ($issue in $issues) { '[{0}] {1}：{2}' -f $issue.Stage, $issue.Path, $issue.Message }
        if ($reportWritten) { "统计报告：$reportPath" }
    }
} finally {
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
