# FY4B 云图自动下载与 IDM 导入工具

本项目提供了一套高度可配置的自动化工具，用于从中国国家气象中心（NMC）下载风云四号B星（FY4B）的云图数据。

项目支持两种主要工作模式：
1. **直接下载模式** ([fy4b_direct_download.py](./fy4b_direct_download.py))：使用 Python 内置库直接下载，支持临时目录缓存、自动移动文件和预下载下一批次。
2. **IDM 导入模式** ([fy4b_idm_import.py](./fy4b_idm_import.py))：将链接导入 Internet Download Manager (IDM) 进行多线程加速下载，并自动管理链接有效期。

所有路径、参数和逻辑均通过 [fy4b_config.json](./fy4b_config.json) 配置，无需修改 Python 源码即可适配不同环境。

## 📁 项目结构

```text
FY4B-auto-download/
├── fy4b_config.json          # [核心] 全局配置文件，修改此处即可调整行为
├── fy4b_direct_download.py   # 主脚本：Python 直接下载模式
├── fy4b_idm_import.py        # 主脚本：IDM 导入模式
├── check_excel.py            # 辅助模块：检查文件状态
├── create_download_links.py  # 辅助模块：根据模板生成下载链接
├── url_template.md           # 下载链接 URL 模板
└── README.md                 # 项目说明文档
```

## ⚙️ 环境要求

### 软件依赖
- **Python 3.10+** (使用了类型提示 `|` 语法)
- **Microsoft Excel**：用于刷新链接有效期（两个主脚本均依赖 Excel COM 接口）。
- **Internet Download Manager (IDM)**：仅在使用 [fy4b_idm_import.py](./fy4b_idm_import.py) 时需要。

### Python 库依赖
```bash
pip install pywin32
```
> **注意**：`pywin32` 是 Windows 平台操作 Excel 和调用外部程序的核心依赖。

## 🛠️ 配置说明 (`fy4b_config.json`)

在运行脚本前，请编辑 [fy4b_config.json](./fy4b_config.json) 以匹配你的本地环境。

### 1. 路径配置 (`路径`)
所有相对路径均相对于脚本所在目录解析。

| 字段 | 说明 | 示例 |
| :--- | :--- | :--- |
| `临时目录` | 下载时的暂存文件夹，下载完成后会自动移至目标目录 | `./云图/FY4B_40H` |
| `目标目录` | 最终保存图片的文件夹 | `./云图/FY4B` |
| `txt文件目录` | 存放 `.txt` 链接文件的目录 | [./](./fy4b_config.json) |
| `Excel文件` | 包含链接生成公式的 Excel 文件路径 | `./链接生成-有效期24小时.xlsx` |
| `模板文件` | URL 模板文件路径 | `./url_template.md` |
| `IDM程序` | IDM 可执行文件的绝对路径 | `C:\Program Files\...\IDMan.exe` |

### 2. Excel 结构配置 (`Excel`)
如果你的 Excel 模板结构与默认不同，请调整以下字段：
- `工作表索引`: 第几个工作表（1 表示第一个）。
- `B162单元格` / `B2单元格`: 用于触发刷新的源和目标单元格。
- `B1单元格`: 存储当前批次过期时间的单元格。
- `链接列`: 下载链接所在的列（如 "E"）。
- `链接起始行` / `链接结束行`: 链接所在的行范围。

### 3. 下载参数 (`下载`)
- `超时秒数`: 单个文件下载超时时间。
- `每N个请求暂停`: 防封禁策略，每下载 N 个文件暂停一次。
- `暂停秒数`: 暂停的时长。
- `图片后缀`: 允许下载的文件扩展名列表。

### 4. 链接生成配置 (`txt文件`)
- `前缀`: 生成的 txt 文件名前缀（如 `截止T：`）。
- `日期格式`: 文件名中的时间格式。
- `链接数量`: 每次生成的链接总数（默认 160 个，对应 40 小时数据）。
- `时间偏移小时`: 起始时间相对于当前时间的偏移（默认回溯 30 小时）。
- `链接间隔分钟`: 卫星云图的时间间隔（网站默认配置为 15 分钟，修改后将无法全量下载）。
- `截止时间偏移分钟`: 用于计算文件名中“截止时间”的偏移量（需≥15分钟）。

## 🚀 使用方法

### 方式一：直接下载（推荐）
适合大多数用户，全自动管理下载、去重、文件移动和链接更新。

```bash
python fy4b_direct_download.py
```

**工作流程：**
1. **检查状态**：检查 Excel 和 `.txt` 文件是否存在。
   - 若都不存在，根据 [url_template.md](./url_template.md) 自动生成新的 `.txt` 文件。
2. **判断过期**：
   - **未过期**：直接读取现有 `.txt`，下载到 `临时目录`（跳过已存在文件）。
   - **已过期**：
     1. 先下载旧 `.txt` 中的剩余链接到 `临时目录`。
     2. 更新 Excel（触发公式重算）或生成新 `.txt`。
     3. 将 `临时目录` 中的所有图片移动到 `目标目录`。
     4. 删除旧 `.txt`，生成新 `.txt`。
     5. **预下载**：立即开始下载新批次的链接到 `临时目录`，为下次运行做准备。

### 方式二：IDM 导入模式
适合拥有 IDM 授权且希望利用其多线程加速的用户。

```bash
python fy4b_idm_import.py
```

**工作流程：**
1. **检查状态**：同直接下载模式，若缺失文件则自动生成。
2. **判断过期**：
   - 若链接未过期，跳过导入。
   - 若链接已过期（或刚生成），调用 IDM 命令行接口导入 `.txt` 中的链接。
3. **更新链接**：
   - 若 Excel 存在：更新 Excel 并重命名 `.txt` 文件。
   - 若 Excel 不存在：根据模板生成新的 `.txt` 文件并删除旧的。

### 方式三：辅助工具

#### 1. 检查文件状态
```bash
python check_excel.py
```
输出当前 Excel 和 `.txt` 文件的存在情况及过期时间，用于调试。

#### 2. 手动生成链接文件
```bash
python create_download_links.py
```
根据配置和模板，手动生成一个新的 `截止T：....txt` 文件，不执行下载操作。

## 📝 链接格式参考

下载链接基于 [url_template.md](./url_template.md) 中的模板生成：

```text
https://image.nmc.cn/product/YYYY/MM/DD/WXBL/SEVP_NSMC_WXBL_FY4B_ETCC_ACHN_LNO_PY_YYYYMMDDhhmm00000.JPG
```

- `YYYY`: 年份
- `MM`: 月份
- `DD`: 日期
- `hh`: 小时
- `mm`: 分钟（00, 15, 30, 45）

## ⚠️ 注意事项

1. **Excel 占用问题**：
   - 脚本运行时会独占打开 Excel 文件。请确保 `链接生成-有效期24小时.xlsx` **未被其他 Excel 进程打开**，否则会导致写入失败。
   
2. **临时目录机制**：
   - [fy4b_direct_download.py](./fy4b_direct_download.py) 使用“先下载到临时目录，再移动”的策略。这确保了只有完整下载的文件才会进入最终目录，避免产生损坏的图片文件。

3. **网络友好性**：
   - 脚本内置了延时机制（`每N个请求暂停`），以避免对服务器造成过大压力或被 IP 封禁。请勿随意将暂停时间设置为 0。

4. **配置文件备份**：
   - 建议定期备份 [fy4b_config.json](./fy4b_config.json) 和 Excel 模板文件，以防配置丢失。

## 🐛 常见问题排查

- **报错 "No module named 'win32com'"**：
  请运行 `pip install pywin32`。如果安装后仍报错，请以管理员身份运行 CMD 并执行 `python Scripts/pywin32_postinstall.py -install`。

- **报错 "Excel Update Failed"**：
  1. 检查 Excel 文件是否被其他程序占用。
  2. 检查 [fy4b_config.json](./fy4b_config.json) 中的单元格引用（如 `B162`, `E2` 等）是否与实际 Excel 文件一致。
  3. 确保 Excel 文件中包含正确的公式，能够根据 `B2` 的变化自动更新 `B1` 和 `E` 列的链接。

- **IDM 导入无反应**：
  1. 确认 [fy4b_config.json](./fy4b_config.json) 中 `IDM程序` 路径正确。
  2. 尝试在 CMD 中手动运行：`"C:\Path\To\IDMan.exe" /s /import "C:\Path\To\links.txt"` 测试是否正常。

- **生成的链接无法下载**：
  1. 检查 [url_template.md](./url_template.md) 中的模板是否与 NMC 最新格式一致。
  2. 检查系统时间是否准确，因为链接生成依赖于当前时间计算。

## 📄 免责声明

本项目仅供个人学习和研究使用。请下载和使用数据时遵守中国国家气象中心及相关数据提供方的使用政策和版权规定。严禁将本工具用于商业目的或对服务器进行恶意攻击。