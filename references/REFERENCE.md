# 调用与维护参考

仅在直接调用、排障或修改 skill 时阅读。

## 参数

以 SKILL.md 所在目录为工作目录：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<真实 msg\file 完整路径>" -AsJson
```

| 参数 | 用途 |
|---|---|
| `-Apply` | 正式移动；省略则只读预览，不创建目录或报告 |
| `-Deduplicate` | 分析重复；与 `-Apply` 合用才回收重复副本 |
| `-MigrateLegacyAudio` | 额外将“整理归纳／其他”中的历史音频迁入“音频”；默认关闭，支持预览，不是只处理旧音频 |
| `-IncludeRecords` | 输出逐文件明细；默认仅汇总和全部失败详情 |
| `-AsJson` | 输出 JSON；省略则为可读文本 |
| `-ListAccounts` | 只读发现账号，不能和 RootPath、Apply、Deduplicate、MigrateLegacyAudio 合用 |
| `-WechatBasePath` | 仅用于账号发现，指定一个 xwechat_files 根目录替代默认位置 |

```powershell
# 发现候选账号
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -ListAccounts -AsJson
# 可选旧音频迁移预览；确认迁移后在明确授权的调用中加 -Apply
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<真实目录>" -MigrateLegacyAudio -AsJson
```

默认仅在以下两个位置发现 `<账号>\msg\file`，不扫描下载文件内容：

1. `C:\Users\<用户名>\xwechat_files`
2. `C:\Users\<用户名>\Documents\xwechat_files`

真实用户目录优先取环境变量 `USERPROFILE`，为空或属于 `CodexSandbox*` 时回退到系统 UserProfile；两者均排除 `CodexSandbox*` 沙箱账号。无法识别时停止并提示指定目录，不枚举其他用户。`Documents` 固定为该用户目录下的子目录，不使用系统重定向到 OneDrive 等位置的文档目录。显式指定 `-WechatBasePath` 仍可替代默认搜索位置。

`Candidates` 给出账号名与完整路径，`SearchRoots` 始终为数组并去重。`SelectionRequired=false` 表示仅有一个候选，不表示已经授权移动；执行仍须显式传入 `-RootPath`。

## 输出接口

| 字段 | 含义 |
|---|---|
| `Run.Phase` | Planned 为预览，Completed 为正式处理结束；是否完整另看 Complete |
| `Run.Statistics` | 本次各分类的计划量或成功移动量；每行 Folder、FileCount、Size |
| `Run.TotalFiles` / `Run.TotalSize` | 上述本次处理总计，不扣除随后回收的副本 |
| `Run.Complete` | 本次扫描与移动是否均无失败 |
| `Run.AudioMigrationCount` | 本次计划或成功迁移的旧音频数量，是本次总数的子集 |
| `Inventory` | 预览为 null；正式返回 Statistics、TotalFiles、TotalSize、Complete，表示最终库存及其扫描完整性 |
| `FileCount` | 本次扫描到的待处理文件数，含后来移动失败的文件 |
| `CategoryCounts` | 八类计划／成功移动数量，包含零计数键 |
| `RenamedCount` / `FailedCount` | 本次计划或成功重命名数／移动失败数 |
| `DuplicateCount` / `DuplicateSize` | 发现的重复副本数量／大小 |
| `RecycledCount` / `RecycledSize` / `RecycleFailedCount` | 实际回收数、大小、回收失败数 |
| `DuplicateCheckComplete` | 未请求去重为 null，否则表示扫描和重复校验是否完整；回收失败另看 RecycleFailedCount |
| `Success` / `Failures` | 整体是否无失败，以及带 Stage、Path、Message 的失败详情 |
| `ReportPath` | 本次成功生成报告的真实完整路径；未生成时 null。回复时直接显示该路径，不提供下载链接，不复制报告到 outputs 等交付目录 |
| `Records` | 仅 IncludeRecords 返回；含 Planned、Moved、Failed、DuplicatePlanned、Recycled、RecycleFailed；KeptPath 指出保留文件 |

迁移旧调用方：`SearchRoot` 改用 `SearchRoots`；`SizeKB` 改用带单位的 `Size`。原顶层 `Statistics`、`TotalFiles`、`TotalSize`、`StatisticsComplete` 已移除；按用途选择 `Run` 或 `Inventory`，完整性字段改为相应对象中的 `Complete`。两者不相互替代；不要把 Inventory=null 当成零库存。

## 分类与统计规则

按扩展名匹配，忽略大小写，不解压文件，不修改文件内容。

| 分类目录 | 扩展名 |
| --- | --- |
| `word` | `.doc` `.docx` `.docm` `.dot` `.dotx` `.dotm` `.wps` `.wpt` |
| `excel` | `.xls` `.xlsx` `.xlsm` `.xlsb` `.xlt` `.xltx` `.xltm` `.csv` `.et` `.ett` |
| `ppt` | `.ppt` `.pptx` `.pptm` `.pot` `.potx` `.potm` `.pps` `.ppsx` `.ppsm` `.dps` `.dpt` |
| `压缩包` | `.zip` `.rar` `.7z` `.tar` `.gz` `.tgz` `.bz2` `.xz` `.zst` `.cab` |
| `DWG` | `.dwg` |
| `PDF` | `.pdf`（忽略扩展名大小写） |
| `音频` | .mp3 .wav .flac .m4a .m4b .aac .ogg .opus .wma .amr .aif .aiff .ape .alac .mid .midi .silk |
| `其他` | 其余文件，包括图片、YML 和无扩展名文件 |

新文件扫描排除“整理归纳”，跳过链接。正式处理后仅非递归删除物理为空的内置分类目录；有子项就保留，不删自建目录、月份目录和历史 CSV，不生成 CSV 日志。

库存统计扫描“整理归纳”下所有直接子文件夹中的递归文件，包含历史和自建分类，排除链接与目标根目录文件。无文件行不显示，零字节文件仍计数。各行与总计独立从原始字节按 1024 换算为 KB、MB、GB，保留两位小数。

去重按分类和大小分组，再比 SHA-256；不同名同内容也算重复。同组按已有文件优先、创建时间、完整路径排序保留一份；显式迁移的旧音频仍算已有文件。默认不迁移时，历史音频仍归属于当前所在分类，不跨分类去重。

## 排障与开发

- 所有 ps1 保持 UTF-8 BOM，兼容 PowerShell 5.1 中文源码。报告同样为 UTF-8 BOM；对齐排版和最终回复全文展示保留。
- 回收实现位于 `scripts/deduplicate.ps1` 和 `scripts/recycle-native.cs`：删除前复核内容与文件状态，锁定保留文件，用 IFileOperation 回调拒绝不满足回收条件的删除。不做永久删除回退。
- Mutex 防止同一根目录同时运行；报告用临时文件替换。修改时复用现有实现，不另写一套分类或删除逻辑。
- 遇到禁止执行脚本或 UnauthorizedAccess，可只读检查 `Get-ExecutionPolicy -List` 和 `Get-ExecutionPolicy`。执行策略与文件权限相互独立，不自动添加 Bypass、解除来源阻止或改写命令绕过限制。需要策略调整时解释具体影响并取得对应授权，遵守组织策略。
- 自动审批拒绝时报告原因，不反复改写同一受限动作。Windows PowerShell 被阻止时中文错误可能乱码，仍可据 UnauthorizedAccess 识别。

开发测试需要 Python 3，日常不需要。在包目录运行 `python ./tests/test_organizer.py`；用环境变量 `WECHAT_TEST_POWERSHELL` 指定 PowerShell 7 或完整路径。测试不放宽执行策略。

测试只使用包父目录下生成的临时文件；回收用例通过 `tests/restore-test-recycle.ps1` 只恢复本次测试文件，恢复成功后清理临时目录。应在可写的隔离副本中运行。修改统计和迁移时优先验证预览不改文件、同名不覆盖、失败保留、重复运行、统计口径和旧音频开关；不扩展为全面安全扫描。

官方参考：[Claude Code skills](https://code.claude.com/docs/en/skills#control-who-invokes-a-skill)、[PowerShell 执行策略](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1)。Claude 的 disable-model-invocation 与 agents/openai.yaml 中的 Codex 配置分别生效；保留现有 skill 名称和显式调用入口。
