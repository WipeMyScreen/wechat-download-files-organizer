param([Parameter(Mandatory=$true)][string]$TestRoot)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$absolute = [IO.Path]::GetFullPath($TestRoot).TrimEnd('\')
if ([IO.Path]::GetFileName($absolute) -notlike 'wechat-organizer-test-*') { throw 'Only generated test roots may be restored.' }
$shell = New-Object -ComObject Shell.Application
$expected = [Collections.Generic.List[string]]::new()
foreach ($item in @($shell.NameSpace(10).Items())) {
    $from = [string]$item.ExtendedProperty('System.Recycle.DeletedFrom')
    if ($from -and ($from.Equals($absolute, [StringComparison]::OrdinalIgnoreCase) -or $from.StartsWith($absolute + '\', [StringComparison]::OrdinalIgnoreCase))) {
        $name = [string]$item.Name
        $expected.Add((Join-Path $from $name))
        $item.InvokeVerb('undelete')
    }
}
for ($attempt = 0; $attempt -lt 50; $attempt++) {
    $remaining = @($expected | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($remaining.Count -eq 0) { break }
    Start-Sleep -Milliseconds 100
}
if (@($expected | Where-Object { -not (Test-Path -LiteralPath $_) }).Count) { throw 'Restoration of generated test items was not confirmed.' }
'Restored only files belonging to this exact generated test root.'
