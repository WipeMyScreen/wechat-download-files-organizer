# Native code is loaded only when an actual recycle operation is needed.
$script:RecycleSourcePath = Join-Path $PSScriptRoot 'recycle-native.cs'

function New-DuplicateEntry($File, [string]$Category, [bool]$Existing) {
    [pscustomobject]@{
        Path = $File.FullName; Category = $Category; Existing = $Existing
        Length = [long]$File.Length; CreationTicks = $File.CreationTimeUtc.Ticks
        WriteTicks = $File.LastWriteTimeUtc.Ticks; Hash = $null
    }
}

function Get-StableFileHash($Entry) {
    Assert-PlainDirectory ([IO.Path]::GetDirectoryName($Entry.Path)) -CheckParents
    $info = Get-Item -LiteralPath $Entry.Path -Force
    if ($info.PSIsContainer -or ($info.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw '只允许比较普通文件。' }
    if ($info.Length -ne $Entry.Length -or $info.CreationTimeUtc.Ticks -ne $Entry.CreationTicks -or $info.LastWriteTimeUtc.Ticks -ne $Entry.WriteTicks) {
        throw '文件在扫描后发生变化，保留该文件。'
    }
    $stream = [IO.File]::Open($Entry.Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $hash = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
        finally { $sha.Dispose() }
        $info.Refresh()
        if ($info.Length -ne $Entry.Length -or $info.LastWriteTimeUtc.Ticks -ne $Entry.WriteTicks) { throw '校验期间文件发生变化，保留该文件。' }
        return $hash
    } finally { $stream.Dispose() }
}

function Get-DuplicatePlan([object[]]$Entries) {
    $pairs = [Collections.Generic.List[object]]::new()
    foreach ($sizeGroup in @($Entries | Group-Object Category, Length | Where-Object Count -gt 1)) {
        $valid = [Collections.Generic.List[object]]::new()
        foreach ($entry in $sizeGroup.Group) {
            try { $entry.Hash = Get-StableFileHash $entry; $valid.Add($entry) }
            catch { Add-Issue 'Duplicate' $entry.Path $_.Exception.Message }
        }
        foreach ($hashGroup in @($valid | Group-Object Hash | Where-Object Count -gt 1)) {
            $ordered = @($hashGroup.Group | Sort-Object @{Expression={ if ($_.Existing) { 0 } else { 1 } }}, CreationTicks, Path)
            $keeper = $ordered[0]
            foreach ($duplicate in $ordered[1..($ordered.Count - 1)]) {
                $pairs.Add([pscustomobject]@{ Keeper = $keeper; Candidate = $duplicate })
            }
        }
    }
    $pairs.ToArray()
}

function Invoke-NativeRecycle($Pair) {
    if (-not ('WechatOrganizerRecycleV1.NativeRecycle' -as [type])) {
        Add-Type -Path $script:RecycleSourcePath -ErrorAction Stop
    }
    $candidate = $Pair.Candidate
    [WechatOrganizerRecycleV1.NativeRecycle]::RecycleDuplicate(
        $candidate.Path, $Pair.Keeper.Path, $candidate.Hash, $candidate.Length,
        $candidate.CreationTicks, $candidate.WriteTicks)
}

function Invoke-DuplicateRecycle($Pair, [scriptblock]$RecycleAction = { param($value) Invoke-NativeRecycle $value }) {
    # Revalidate both sides even when an alternate backend is used by isolated tests.
    $candidateHash = Get-StableFileHash $Pair.Candidate
    $keeperHash = Get-StableFileHash $Pair.Keeper
    if ($candidateHash -ne $Pair.Candidate.Hash -or $keeperHash -ne $Pair.Candidate.Hash) {
        throw '重复文件或保留文件已变化，保留该组文件。'
    }
    $response = & $RecycleAction $Pair
    if ($null -eq $response -or -not $response.Recycled) {
        $reason = 'Windows 回收站接口未确认成功，文件保留；不会永久删除。'
        if ($null -ne $response -and $response.Error) { $reason = $response.Error }
        throw $reason
    }
    return $response
}
