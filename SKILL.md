---
name: wechat_download_files_organizer
description: 整理本机微信下载目录中的文件，按 Word、Excel、PPT、压缩包、DWG、PDF 和其他分类移动，并生成文件数量及自动换算单位的大小统计。当用户要求“整理微信文件”“归纳微信下载文件”或预览微信文件分类时使用。
---

# 微信文件整理

使用本 skill 自带的 [scripts/organize.ps1](scripts/organize.ps1)，不要临时重写分类或移动逻辑。适用于 Windows 本机的 Codex 和 Claude Code；使用 Windows 自带的 PowerShell 5.1，也兼容 PowerShell 7，无需额外运行时。

脚本及 TXT 报告保持 UTF-8 with BOM，以兼容 PowerShell 5.1 中文。遵守现有执行策略及工具权限，不擅自调整策略或绕过限制；遇到限制时报告原因，按 [README 排障说明](README.md#执行策略与自动模式权限) 处理。

## 选择本次目录

默认检查当前 Windows 用户目录和 Windows 文档目录下的 `xwechat_files\<账号文件夹>\msg\file`，汇总两处找到的账号，不固定账号。文档目录使用系统实际位置（包括重定向位置），两处路径相同时只检查一次；显式提供 `-WechatBasePath` 时只查指定位置。

1. 用户已在本次请求中给出完整文件目录时直接使用；给出账号文件夹名时，先发现候选目录核对，唯一匹配后使用，无需重复询问。
2. 否则执行下方账号发现命令，列出候选账号名及完整路径，询问并等待用户选择。**只有一个候选也必须询问。** 同名账号以完整路径区分。不根据当前工作目录、排序、修改时间或上次选择自行确定账号，不推测登录状态。
3. 用户可回答账号名、序号或完整路径。没有候选时询问实际的 `msg\file` 路径；也可对用户给出的自定义位置加 `-WechatBasePath "<xwechat_files 路径>"` 重新发现，不创建猜测的目录。
4. 同一次任务内沿用选定目录，不处理其他账号。取消或尚未选择时不整理、不创建目录和报告。

发现阶段仅检查账号目录结构，跳过不含 `msg\file` 的目录及链接，不遍历账号文件。预览和整理必须显式传入 `-RootPath`。

## 确认本次操作

目录确定后（包括用户直接提供完整路径时），再次询问并等待用户选择：

> 已选择 `<账号名或完整路径>`，本次要进行“只读预览”（不移动、不写报告），还是“正式移动”（分类移动并更新统计报告）？

- 选择“只读预览”：运行不带 `-Apply` 的命令，汇报后结束，不自动接着移动。用户随后明确要求移动同一目录时，再执行 `-Apply`，无需重复询问。
- 选择“正式移动”：运行带 `-Apply` 的命令。
- 尚未回答或取消：不运行预览或移动命令。

确认账号只确定操作目录，不代表选择了移动；最初的“整理”请求也不能省略这次操作选择。

## 调用命令

以下命令以本 `SKILL.md` 所在目录为工作目录。将占位符替换为本次选定的真实路径，路径须加引号。其他终端的调用方式见 [README](README.md#在-claude-code-中使用)。

发现账号：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -ListAccounts -AsJson
```

只读预览，不移动文件、不创建目录或报告：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<已选定的完整文件目录>" -AsJson
```

正式整理：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<已选定的完整文件目录>" -Apply -AsJson
```

根据用户在上述确认步骤中选择的操作调用命令。安装或修改 skill 不授权整理现有文件。

默认返回汇总及失败详情。需要逐文件明细时加 `-IncludeRecords`，例如：

```powershell
powershell.exe -NoProfile -NonInteractive -File "./scripts/organize.ps1" -RootPath "<已选定的完整文件目录>" -AsJson -IncludeRecords
```

`-IncludeRecords` 也适用于正式整理；省略 `-AsJson` 输出可读文本。默认 JSON 不含 `Records`，默认文本不列出逐文件移动路径。需要旧版完整 JSON 明细的调用方添加 `-IncludeRecords`。

## 整理与统计

- 递归扫描源目录，排除整个 `整理归纳`，跳过文件和目录链接；按扩展名分类到源目录下的 `整理归纳`，不保留月份层级。
- 保留原空目录，不解压、不改文件内容。同名追加 ` (1)`、` (2)`，不覆盖；单个移动失败保留源文件、记录原因并继续，不自动重试。

| 分类 | 扩展名 |
|---|---|
| word | .doc .docx .docm .dot .dotx .dotm .wps .wpt |
| excel | .xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .et .ett |
| ppt | .ppt .pptx .pptm .pot .potx .potm .pps .ppsx .ppsm .dps .dpt |
| 压缩包 | .zip .rar .7z .tar .gz .tgz .bz2 .xz .zst .cab |
| DWG | .dwg |
| PDF | .pdf |
| 其他 | 其余文件，包括图片、YML 和无扩展名文件 |

扩展名忽略大小写。每次正式执行均覆盖 `整理归纳\整理统计.txt`，包括零移动和部分失败时。不生成移动记录 CSV，也不删除历史 CSV。

统计 `整理归纳` 下所有直接子文件夹的递归文件数和逻辑大小，包含已有文件及用户新增目录；跳过链接，空目录显示零，目标根目录中的文件不计入总计。各行及总计基于原始字节数独立选择 KB、MB、GB，按 1024 换算，保留两位小数，空目录显示 `0.00 KB`。时间采用北京时间。

报告由脚本生成中英文列宽对齐的三列表格，完整样例见 [README](README.md#整理结果与统计报告)。JSON 的 `Statistics[].Size` 含单位；保留 `SizeKB` 两位小数字符串以兼容已有调用。

## 解释结果

预览时使用 `Statistics` 输出“分类 / 文件数量 / 文件大小”三列表格，并用 `TotalFiles`、`TotalSize` 显示总计。大小仅包含本次待整理的文件，不含 `整理归纳` 中已有文件；七类均列出，空分类显示 `0.00 KB`。直接使用脚本返回的 `Size`，自动按 1024 换算为 KB、MB 或 GB，保留两位小数；总计基于原始字节数独立换算。

正式整理时汇报各分类移动数及报告位置；此时 `Statistics` 和总计表示目标目录整理后的实际统计。两种模式均汇报重命名数、失败和跳过链接情况。默认保留 `Failures` 全部失败详情；完整移动记录仅在 `-IncludeRecords` 时返回，不另写日志。

`Success=false` 时解释失败；`StatisticsComplete=false` 时说明统计不完整；报告写入失败也必须告知。文件移动后，微信聊天记录中的文件入口可能无法按原路径打开，可从分类目录打开。
