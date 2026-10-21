---
name: wechat_download_files_organizer
description: 整理本机微信下载目录中的文件，按 Word、Excel、PPT、压缩包、DWG 和其他分类移动，并生成文件数量及自动换算单位的大小统计。当用户要求“整理微信文件”“归纳微信下载文件”或预览微信文件分类时使用。
---

# 微信文件整理

微信账号目录因登录用户而异，不能固定为某一个账号。默认在当前 Windows 用户的文档目录下查找 `xwechat_files\<账号文件夹>\msg\file`。
目标为本次用户选定源目录下的 `整理归纳`，分类目录为 `word`、`excel`、`ppt`、`压缩包`、`DWG`、`其他`。

使用本 skill 自带的 [scripts/organize.ps1](scripts/organize.ps1)，通过 Windows 自带的 `powershell.exe`（Windows PowerShell 5.1）执行；也兼容 PowerShell 7。不要临时重写分类或移动逻辑，也不要通过自动化后台运行。

## 运行环境

适用于 Windows 本机的 Codex 和 Claude Code。无需安装 PowerShell 7、Python 或其他运行时，也不依赖 Codex 私有运行时路径。此 skill 使用 Windows 文件系统和北京时间时区，不适用于原生 macOS/Linux 环境。

- 优先使用 `powershell.exe`；只有系统自带的 5.1 也可完整运行。有 PowerShell 7 时也可用 `pwsh` 执行同一脚本。
- PowerShell 命令中的带引号的可执行文件路径前要加 `&`；Git Bash 中直接执行带引号的路径，不加 `&`。路径可能含空格和中文，必须引用。
- 若 Git Bash 的 PATH 找不到 `powershell.exe`，用 `"${SYSTEMROOT:-C:/Windows}/System32/WindowsPowerShell/v1.0/powershell.exe"`；PowerShell 中可用 `& "$env:SystemRoot/System32/WindowsPowerShell/v1.0/powershell.exe"`。定位到可执行文件后直接继续原任务，不要求安装 7。
- 默认命令遵守本机有效执行策略，不附加放宽执行策略的参数。版本兼容不代表系统允许执行脚本。
- 出现 `UnauthorizedAccess`、`PSSecurityException` 或“禁止运行脚本”时，运行只读诊断 `powershell.exe -NoProfile -NonInteractive -Command 'Get-ExecutionPolicy -List; Get-ExecutionPolicy'`，说明有效策略和作用域，并停止本次脚本调用。
- 普通“整理文件”请求不授权修改执行策略。用户明确同意前，不运行 `Set-ExecutionPolicy`，不附加 `-ExecutionPolicy Bypass`、`Unrestricted` 或其他放宽策略的参数，也不改用内联脚本、编码命令或其他解释器来避开已知限制。
- 若是用户自行管理的电脑，可解释 `RemoteSigned` 允许本地未签名脚本运行，对带 Internet 来源标记的脚本要求可信签名。用户可明确选择仅当前进程的临时授权，或 `CurrentUser` 范围的持久配置；只有取得对应授权后才执行准确范围的调整。组织策略要求禁止执行或签名时，由管理员处理，不自动移除来源标记。
- Claude 自动模式检查无结果或拒绝时，说明拒绝的具体动作和原因，不用 `sleep` 或反复改写同一受限命令重试，也不自动放宽 Claude 权限。脚本运行权限与整理文件授权是两件事。
- `scripts/organize.ps1` 必须保持 **UTF-8 with BOM**，保证 5.1 正确读取中文。脚本已固定控制台输出为 UTF-8，TXT 报告也使用 UTF-8 with BOM。
- 如果当前工具禁止读取或执行 skill 路径，说明受限路径及需要的访问权限；不得据此推断脚本不兼容，也不临时重写脚本。

## 调用

以下命令中的 `./scripts/organize.ps1` 均相对于本 `SKILL.md` 所在目录。执行前，根据当前加载的 skill 路径确定该目录，并将命令工具的工作目录设为该目录；待整理的微信文件目录仍通过 `-RootPath` 单独传入。

每次用户点名调用本 skill 发起整理或预览时，先确定本次要操作的目录：

1. 若用户已在本次请求中明确给出完整文件目录，直接使用该目录。若给出账号文件夹名，先列出候选目录核对对应路径；唯一匹配时使用该目录，无需重复询问。
2. 否则，先执行只读账号发现命令：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -ListAccounts -AsJson
```

3. 把候选账号文件夹名及完整路径列给用户，询问本次选择哪个，并等待回答；只有一个候选时也要询问。不得根据当前工作目录、目录顺序、修改时间或上一次选择自行确定账号；不推测哪个账号正在登录。
4. 用户可回答账号文件夹名、候选序号或完整路径。没有候选时，询问实际的 `msg\file` 完整路径；也可对用户指定的 `xwechat_files` 位置运行 `-ListAccounts -WechatBasePath "<指定位置>" -AsJson` 重新查找。不要创建猜测的账号目录。
5. 选择完成后，同一次任务内保持该目录，正式整理直接执行；不再重复询问授权。不批量处理其他账号。用户取消或尚未作答时，不整理、不创建分类目录和报告。

发现阶段只检查账号目录结构，跳过 `all_users`、`Backup` 等不含 `msg\file` 的目录以及链接，不遍历候选账号中的文件。整理脚本不再有默认账号，缺少 `-RootPath` 会报错停止。

用户要求预览时，执行只读模式（默认），不创建目录或报告：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<本次已选定的完整文件目录>"
```

用户要求整理时，该请求已授权实际移动，直接执行，不重复请求确认：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<本次已选定的完整文件目录>" -Apply
```

执行命令前将示例占位符替换为用户本次选定的真实路径。可加 `-AsJson` 读取结构化结果。只创建/安装/修改 skill 的请求不授权整理现有文件。

## 行为

- 遍历源根目录和所有子目录，排除整个 `整理归纳`，跳过文件/目录链接；路径校验失败时停止，不绕过保护。
- 文件按下表扩展名（忽略大小写）直接归类，不保留月份层级；保留原空目录，不解压、不改文件内容。
- 同名保留全部，添加 ` (1)`、` (2)`，绝不覆盖。单个移动失败保留源文件并继续。

| 分类 | 扩展名 |
|---|---|
| word | .doc .docx .docm .dot .dotx .dotm .wps .wpt |
| excel | .xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .et .ett |
| ppt | .ppt .pptx .pptm .pot .potx .potm .pps .ppsx .ppsm .dps .dpt |
| 压缩包 | .zip .rar .7z .tar .gz .tgz .bz2 .xz .zst .cab |
| DWG | .dwg |
| 其他 | 其余，包括 PDF、图片、YML 和无扩展名文件 |

每次正式执行均在目标根目录覆盖 UTF-8 的 `整理统计.txt`，包括零移动和部分失败时。统计所有直接子文件夹（含用户新增目录）的递归文件数与文件逻辑大小，空目录显示零；跳过链接。各文件夹和总计分别按真实字节数自动选择 **KB、MB 或 GB**：小于 1 MB 使用 KB，达到 1 MB 且小于 1 GB 使用 MB，达到 1 GB 使用 GB。按 **1 KB = 1024 字节、1 MB = 1024 KB、1 GB = 1024 MB** 换算，保留两位小数，不列精确字节数，空目录显示 `0.00 KB`。总计先累加原始大小，再独立选择单位，不累加各行已舍入的显示值。目标根目录中的文件不计入总计。时间采用北京时间。

JSON 结果的 `Statistics[].Size` 为含单位的显示文本（如 `1.50 GB`）；原有 `SizeKB` 字段保留为以 KB 表示的两位小数字符串，供已有调用兼容使用。

正式执行只生成或更新 `整理统计.txt`，不生成 CSV 移动记录，也不删除已有的历史 CSV。收到的普通 CSV 文件仍按上表归类到 excel。若读取目录失败导致统计不完整，报告明确标注，不能把不完整统计称为完整结果。

执行后汇报各分类移动数（预览时为预计数）、重命名数、失败/跳过链接情况及报告位置。移动明细与失败原因仅通过运行结果返回，不另写日志文件。结果 `Success=false` 时解释失败；统计报告写入失败也必须告知，不能仅报告移动成功。文件移动后微信原聊天记录中的文件入口可能无法通过原路径打开，用户可从分类目录打开。

## TXT 统计报告格式

`整理统计.txt` 使用以下格式，数据仅作示例。实际输出时替换为本次北京时间和真实统计数据；按前述统计范围列出全部直接子文件夹（含用户新增目录），并保持列对齐、数量单位为“个”、大小单位自动选择 KB/MB/GB 且保留两位小数。统计不完整时仍须明确标注。

```text
微信文件整理统计
============================================
统计时间：2026-10-21 15:30:00（北京时间）

文件夹         文件数量         文件大小
--------------------------------------------
word            12 个           18.00 MB
excel            8 个            6.00 MB
ppt              5 个           45.00 MB
压缩包           4 个          120.00 MB
DWG              7 个           32.00 MB
其他            15 个           24.00 MB
--------------------------------------------
总计            51 个          245.00 MB

说明：
统计包含子文件夹中的文件。
本报告及移动记录 CSV 不计入统计。
============================================
```

移动记录 CSV 指目标根目录中已有的历史移动记录，本次整理不生成新的移动记录 CSV；普通 CSV 文件仍按前述规则归入 excel 并计入统计。
