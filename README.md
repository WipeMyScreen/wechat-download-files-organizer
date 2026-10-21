# wechat_download_files_organizer

一个非常小的skill, 用于整理 Windows 本机微信下载的文件。

在当前 Windows 用户的文档目录下查找 `xwechat_files\<账号文件夹>\msg\file`, 然后把微信下载的文件,移动到\msg\file里的"整理归纳"这个文件夹内的相应目录,  目录为 `word`、`excel`、`ppt`、`压缩包`、`DWG`、`其他`

按文件扩展名将文件归入 Word、Excel、PPT、压缩包、DWG 和其他分类，并生成包含文件数量与 KB 大小的 TXT 统计报告。

支持先预览再整理；同名文件自动编号，保留全部文件。

## 功能

- 查找当前 Windows 用户文档目录下的微信账号文件夹，由用户选择本次要整理的账号；也可直接指定文件目录。
- 递归扫描源目录，按扩展名分类，不保留原有月份或子目录层级。
- 默认只读预览；使用 `-Apply` 才实际移动文件并写入报告。
- 同名文件追加 ` (1)`、` (2)` 等编号，避免覆盖。
- 跳过文件和目录链接，保留源目录中的空文件夹。
- 单个文件移动失败时保留源文件并继续，结果中返回失败原因。
- 统计分类目录及用户新增目录中的文件数量和大小，时间采用北京时间。

## 环境要求

| 环境 | 要求 |
| --- | --- |
| 操作系统 | Windows |
| 执行环境 | PowerShell 7 或更新版本，终端中可运行 `pwsh` |
| Codex | 使用 skill 方式调用时需要；直接运行脚本不需要 Codex |
| 文件权限 | 预览需要读取权限；正式整理需要源目录及目标目录的相应读写权限 |
| Python | 仅运行测试时需要 Python 3，日常整理不需要 |

可以通过以下命令检查 PowerShell 版本：

```powershell
pwsh --version
```

脚本通过本地文件系统操作文件，不需要微信登录接口或 API 密钥。

## 安装到 Codex

1. 从本仓库下载 ZIP 并解压，或使用 Git 克隆仓库。
2. 将包含 `SKILL.md` 的整个文件夹命名为 `wechat_download_files_organizer`。
3. 把该文件夹放入当前用户的 `%USERPROFILE%\.agents\skills\` 目录。目录不存在时先创建；若已有同名 skill，请先备份。

安装后的结构应为：

```text
%USERPROFILE%\.agents\skills\
└── wechat_download_files_organizer\
    ├── README.md
    ├── SKILL.md
    ├── agents\
    │   └── openai.yaml
    ├── scripts\
    │   └── organize.ps1
    └── tests\
        └── test_organizer.py
```

如果 Codex 没有显示新安装的 skill，请重启 Codex 后重试。

## 在 Codex 中使用

此 skill 当前配置为显式调用，请在对话框中按$符号, 主动调用 `$wechat_download_files_organizer`。

### 预览分类

```text
使用 $wechat_download_files_organizer 预览微信下载文件的分类情况，先不要移动文件。
```

### 正式整理

```text
使用 $wechat_download_files_organizer 整理我的微信下载文件，并生成统计报告。
```

未指定目录时，Codex 会列出找到的账号文件夹及完整路径，等待你选择；只有一个候选账号时也会询问。选择后，同一次任务内使用该目录，不批量处理其他账号。

也可以在请求中直接给出要整理的 `msg\file` 完整路径。微信文件保存在自定义位置时，请提供实际路径。

请求“整理”会授权实际移动文件；请求“预览”只读取文件。安装或修改 skill 本身不会触发整理。

## 直接运行脚本

以下命令均在本仓库根目录（即 `SKILL.md` 所在目录）执行。先将路径占位符替换为本机实际目录。

### 1. 列出候选账号

```powershell
pwsh -NoProfile -File "./scripts/organize.ps1" -ListAccounts -AsJson
```

默认搜索当前 Windows 用户文档目录中的 `xwechat_files\<账号文件夹>\msg\file`。发现阶段仅检查账号目录结构，不遍历账号中的文件。

指定其他 `xwechat_files` 位置：

```powershell
pwsh -NoProfile -File "./scripts/organize.ps1" -ListAccounts -WechatBasePath "<xwechat_files 的完整路径>" -AsJson
```

### 2. 只读预览

```powershell
pwsh -NoProfile -File "./scripts/organize.ps1" -RootPath "<选定的 msg\file 完整路径>" -AsJson
```

预览不会移动文件，也不会创建分类目录或统计报告。

### 3. 正式整理

```powershell
pwsh -NoProfile -File "./scripts/organize.ps1" -RootPath "<选定的 msg\file 完整路径>" -Apply -AsJson
```

`-AsJson` 可省略，省略后输出可读文本。`-ListAccounts` 不能与 `-RootPath` 或 `-Apply` 同时使用；`-WechatBasePath` 仅用于列出账号。预览和正式整理都必须明确提供 `-RootPath`。

## 文件分类

按扩展名匹配，忽略大小写，不解压文件，不修改文件内容。

| 分类目录 | 扩展名 |
| --- | --- |
| `word` | `.doc` `.docx` `.docm` `.dot` `.dotx` `.dotm` `.wps` `.wpt` |
| `excel` | `.xls` `.xlsx` `.xlsm` `.xlsb` `.xlt` `.xltx` `.xltm` `.csv` `.et` `.ett` |
| `ppt` | `.ppt` `.pptx` `.pptm` `.pot` `.potx` `.potm` `.pps` `.ppsx` `.ppsm` `.dps` `.dpt` |
| `压缩包` | `.zip` `.rar` `.7z` `.tar` `.gz` `.tgz` `.bz2` `.xz` `.zst` `.cab` |
| `DWG` | `.dwg` |
| `其他` | 其余文件，包括 PDF、图片、YML 和无扩展名文件 |

## 整理结果与统计报告

文件移动到本次选定源目录下的 `整理归纳`：

```text
msg/file/
└── 整理归纳/
    ├── word/
    ├── excel/
    ├── ppt/
    ├── 压缩包/
    ├── DWG/
    ├── 其他/
    └── 整理统计.txt
```

每次正式运行都会覆盖更新 UTF-8 编码的 `整理统计.txt`，即使本次没有文件需要移动。再次整理时会跳过整个 `整理归纳` 目录，避免重复移动已归类文件。

统计范围为 `整理归纳` 下所有直接子文件夹中的递归文件，包含已有文件及用户新增目录；空目录显示零，链接跳过。大小按 **1 KB = 1024 字节** 换算，保留两位小数，使用文件逻辑大小。`整理归纳` 根目录中的文件不计入总计。

下面为报告格式示例，时间和数据仅作演示：

```text
微信文件整理统计
============================================
统计时间：2026-10-21 15:30:00（北京时间）

文件夹         文件数量         文件大小
--------------------------------------------
word            12 个        18432.00 KB
excel            8 个         6144.00 KB
ppt              5 个        46080.00 KB
压缩包           4 个       122880.00 KB
DWG              7 个        32768.00 KB
其他            15 个        24576.00 KB
--------------------------------------------
总计            51 个       250880.00 KB

说明：
统计包含子文件夹中的文件。
本报告及移动记录 CSV 不计入统计。
============================================
```

本次运行不生成移动记录 CSV，也不删除已有历史 CSV。示例中排除的移动记录指 `整理归纳` 根目录中已有的历史记录；普通 CSV 文件仍归入 `excel` 并计入统计。

若目录读取失败，报告会标注统计不完整。请检查运行结果中的失败信息；JSON 结果的 `Success=false` 表示存在失败，报告写入失败也会在结果中反映。

## 使用前须知

- 正式整理会实际移动文件。移动后，微信原聊天记录中的文件入口可能无法按原路径打开，可从 `整理归纳` 中打开文件。
- 分类后不保留原有月份或子目录层级。需要保留原始结构时，请先自行备份。
- 文件被其他程序占用或权限不足时可能移动失败，请检查结果后再处理。

## 项目文件

| 文件 | 用途 |
| --- | --- |
| [SKILL.md](SKILL.md) | Codex 的调用流程、分类规则和输出要求 |
| [agents/openai.yaml](agents/openai.yaml) | 显示信息和显式调用配置 |
| [scripts/organize.ps1](scripts/organize.ps1) | 账号发现、只读预览、文件移动与统计实现 |
| [tests/test_organizer.py](tests/test_organizer.py) | 基于临时文件的集成测试 |

## 运行测试

需要 Windows、PowerShell 7 和 Python 3。在仓库根目录运行：

```powershell
python ./tests/test_organizer.py
```

测试使用独立临时目录，不操作真实微信下载目录。临时目录创建在仓库的父目录下，因此该位置需要写入权限；涉及文件符号链接的测试在缺少相应 Windows 权限时会跳过。
