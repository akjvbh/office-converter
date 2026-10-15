# 办公文档转换器 (Office Converter)

> 一个跑在自己电脑上的办公文档工具箱：**格式转换 · 文件压缩 · PDF 合并 · PDF 可视化编辑 · 扫描件增强 · 英文文献翻译**。
> FastAPI 后端 + 零依赖原生前端，双击即可启动，浏览器打开就是完整界面。

![Python](https://img.shields.io/badge/Python-3.10%2B-3776ab?logo=python&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-0.100%2B-009688?logo=fastapi&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078d4?logo=windows&logoColor=white)
![Frontend](https://img.shields.io/badge/Frontend-Vanilla%20JS-f7df1e?logo=javascript&logoColor=black)
![License](https://img.shields.io/badge/License-MIT-green)

---

## 它解决什么问题

日常办公里最烦的几件事，这里一次做完，**文件不上传第三方服务器**：

- 收到的 `docx / xlsx / pptx` 要转 PDF，但不想买会员、也不想把合同传到别人的云端
- 扫描 App 导出的几十页 PDF 要拆页、删页、改错字、盖章签名
- 手机拍的纸质文件要变成干净的扫描件
- 一堆 PDF 要合成一本，或者按指定页码插进去
- 照片太大发不出去，PDF 太大上传被拒
- 英文文献要读，但一段段复制到翻译网站太慢

---

## 功能一览

| 功能 | 说明 |
|---|---|
| 🔄 **格式转换** | Word / Excel / PPT / PDF / 图片 / Markdown / TXT / CSV 互转，支持批量与整个文件夹 |
| 🗜️ **文件压缩** | 四档压缩级别，选文件后可先「预估」四档压缩后的大小再决定 |
| 📑 **PDF 转 JPG** | 选一个文件夹，自动筛出所有 PDF 逐页导出 JPG，**每张不超过 300KB**，打包 ZIP |
| 📷 **扫描照片** | 手机拍的照片自动找纸张四角 → 裁切 → 透视矫正 → 去阴影 → 增强为扫描件 |
| 🔗 **PDF 合并** | 按顺序合并 / 逐页交替 / 前插 / 后插 / 插入到指定页之后，列表可拖拽排序 |
| ✏️ **PDF 可视化编辑** | 双击原文字直接改写、加字、白底遮盖、高亮划线、矩形圆形、手写签名；页面拖拽排序、复制、删除、旋转 |
| 🌐 **文献翻译** | PDF / DOCX / TXT 逐段调用大模型翻译，输出中英对照、纯译文或**保留原版式的 PDF** |

### 界面特性

- **深色 / 浅色主题**，跟随系统并可手动切换
- **整篇可撤销**：标注、删页、旋转、重排都进同一个撤销栈（`Ctrl+Z` / `Ctrl+Shift+Z`）
- **切换功能不丢文件**：已选文件在选项卡之间保留，可以直接换个方式处理同一批文件
- **批量结果可追溯**：失败文件单独列出，支持一键复制失败清单和错误详情
- 拖拽文件夹、键盘操作（`PageUp/PageDown` 翻页、`Delete` 删除对象、`Ctrl+S` 导出）
- 响应式布局，窄屏（平板/手机）同样可用

---

## 快速开始

### 前置条件

- **Windows 10/11** + **Python 3.10+**
- **Microsoft Office**（Word / Excel / PowerPoint）— *可选*
  仅「Office 文档转 PDF」「图片 OCR 转 Word」依赖它；没有 Office 时其它功能完全正常，
  `docx → PDF` 会自动改用内置排版引擎并给出提示。

### 安装

```powershell
git clone https://github.com/<你的用户名>/office-converter.git
cd office-converter

python -m venv .venv-ocr
.\.venv-ocr\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 运行

```powershell
# 方式 A（最省事）：直接双击 start.bat，或在本目录执行
.\start.bat
```

如果不想用 `start.bat`，可以用虚拟环境的解释器直接启动 ——
**注意不要直接敲 `python main.py`**，原因见下面「Windows 上 `python` 打不开」：

```powershell
# 方式 B：用虚拟环境里的解释器（推荐写全路径，最不会出错）
.\.venv-ocr\Scripts\python.exe main.py
```

或者先激活虚拟环境再运行：

```powershell
# 方式 C：激活虚拟环境
.\.venv-ocr\Scripts\Activate.ps1     # 若报「禁止运行脚本」，见下面的说明
python main.py
```

然后浏览器打开 **http://127.0.0.1:8000**

服务默认监听 `0.0.0.0:8000`，同一局域网内的同事用 `http://<你的内网IP>:8000` 即可访问。

> ⚠️ 服务没有鉴权、没有 HTTPS，设计用途是**本机或可信局域网自用**，请勿直接暴露到公网。

#### Windows 上 `python` 打不开 / 提示 Microsoft Store

如果在命令行敲 `python` 出现这样的提示：

```
Python was not found; run without arguments to install from the Microsoft Store,
or disable this shortcut from Settings > Apps > Advanced app settings > App execution aliases.
```

说明你 PATH 里的 `python` 是 **Windows 自带的“应用执行别名”占位程序**，不是真正的 Python。
虚拟环境本身是好的，只是这个占位程序挡在前面。三种解决办法任选其一：

1. **直接用虚拟环境的解释器**（无需任何系统设置）：
   ```powershell
   .\.venv-ocr\Scripts\python.exe main.py
   ```
2. **关闭别名**：设置 → 应用 → 高级应用设置 → 应用执行别名，把 `python.exe` 和
   `python3.exe` 两个开关都关掉，然后新开一个终端。
3. **用 `py` 启动器**（如果装过 python.org 的安装包，它通常可用）：
   ```powershell
   py -3.11 main.py
   ```

#### PowerShell 报「禁止运行脚本」（Activate.ps1 无法加载）

PowerShell 默认执行策略是 `Restricted`，会拦住 `Activate.ps1`：

```
无法加载文件 ...\Activate.ps1，因为在此系统上禁止运行脚本。
```

要么改用上面的方式 A / B（**不需要激活**，效果完全一样），要么放开当前用户的策略：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

---

## 配置

复制 `.env.example` 为 `.env` 后按需填写（**已存在的真实环境变量优先级更高**）：

```ini
# 文献翻译（唯一需要密钥的功能）
TRANSLATE_API_KEY=sk-xxxxxxxx
TRANSLATE_BASE_URL=https://api.deepseek.com/v1
TRANSLATE_MODEL=deepseek-chat

# 上传限制（MB，必须是大于 0 的整数）
MAX_FILE_SIZE_MB=200
MAX_BATCH_SIZE_MB=500

# Office COM 超时（秒）
OFFICE_TIMEOUT_SECONDS=900
OFFICE_PROBE_TIMEOUT_SECONDS=45
OFFICE_RECHECK_SECONDS=600
```

翻译接口使用 **OpenAI 兼容** 的 `/chat/completions`，所以 DeepSeek、OpenAI、通义千问等都可以直接用，
只需改 `TRANSLATE_BASE_URL` 和 `TRANSLATE_MODEL`。未配置密钥时只有「文献翻译」不可用，其它功能不受影响。

### 术语表

根目录的 [`glossary.txt`](glossary.txt) 是**默认术语表**，每行 `英文=中文`（`#` 开头为注释）：

- 翻译时后端自动携带，界面上的术语表输入框也会自动预填
- 保存后**立即生效，无需重启**
- 界面里填写的内容排在文件内容之后，可以临时覆盖或补充

---

## 支持的转换

| 源格式 | 可转换为 |
|---|---|
| `doc` | pdf, docx, txt |
| `docx` | pdf, txt, html |
| `xls` | pdf, xlsx, csv |
| `xlsx` | pdf, csv |
| `ppt` | pdf, pptx |
| `pptx` | pdf, png |
| `pdf` | docx, png, jpg, txt |
| `csv` | xlsx |
| `md` | html, pdf |
| `txt` | pdf, docx |
| `png` `jpg` `jpeg` `jfif` `bmp` `webp` `gif` `tiff` | 图片互转, pdf, docx(OCR), 部分支持 ico |

`GET /api/formats` 会返回实时的完整映射，前端据此生成下拉选项。

### 关于 Office 转 PDF

Word / Excel / PPT 转 PDF 默认调用本机 Microsoft Office（COM 自动化），排版与 Office 输出一致。

- 每次转换都有超时保护（默认 900 秒），Office 卡死时会中止本次转换并清理新起的 Office 进程
- 首次用到 Office 时会做一次健康探测（默认最多 45 秒）；探测失败后不再重复等待，
  默认 10 分钟后再自动探测一次 —— 因此在任务管理器结束卡死的 Office 进程后，**无需重启服务**即可恢复
- Word 不可用时（Office 更新导致自动化失效、弹出激活窗口、无可交互桌面等），
  `docx → PDF` 会自动改用内置排版引擎，结果区给出「已用内置排版引擎」提示
- 内置引擎无法解析旧版二进制 `.doc`，遇到 `.doc` 请先用 Word 另存为 `.docx`

---

## API

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/api/formats` | 支持的转换映射 |
| `GET` | `/api/config` | 公开配置：上传上限、翻译是否已配置、默认术语表 |
| `POST` | `/api/convert` | 单文件转换。`file`、`target` |
| `POST` | `/api/compress` | 单文件压缩。`file`、`level`（`80`/`60`/`40`/`20`） |
| `POST` | `/api/batch` | 批量转换/压缩。`files[]`、`mode`（`convert`/`compress`）、`target` 或 `level` |
| `POST` | `/api/pdf-folder-to-jpg` | 批量 PDF 逐页转 ≤300KB JPG。`files[]` |
| `POST` | `/api/scan-photo` | 照片裁切矫正并增强为扫描件。`file` |
| `POST` | `/api/pdf-merge` | PDF 合并。`files[]`、`mode`、`insert_index`（**0-based**） |
| `POST` | `/api/pdf-preview` | PDF 页面预览。`file`，最多返回前 50 页 |
| `POST` | `/api/pdf-edit` | PDF 编辑。`file`、`delete_pages`、`rotations`、`page_order`、`annotations` |
| `POST` | `/api/translate` | 提交翻译任务，立即返回 `task_id` |
| `GET` | `/api/translate/progress/{task_id}` | 查询翻译进度，完成后 `result` 内含下载地址 |
| `GET` | `/api/download/{task_id}` | 下载结果 |

<details>
<summary><b>页码与坐标约定（二次开发必读）</b></summary>

这些约定在前后端之间是硬性的，改前端时最容易踩：

- `delete_pages`、`rotations`、`page_order` 里的页码都是**原始 PDF 的 1-based 页码**
- `rotations` 是 `{"页号": 旋转角度}`，角度是**增量**而非绝对值，例如 `{"1":90}`
- `page_order` 是「原始页码」按目标顺序排列的列表，**同时充当白名单**：没列进去的页会被丢弃；重复写同一页码即可复制该页
- 删除与重排是**同时**生效的，`page_order` 引用的始终是原始页码，不受删除影响
- `annotations` 里每个对象的 `page` 是 **0-based**；坐标是**页面左上角原点的 PDF 坐标（点）**。
  标注在重排**之前**写入，因此重排场景下要传「该页在 `page_order` 中的位置」，而不是原始页码
- `insert_index`（合并接口）是 **0-based**：`0` 表示插到首页之前，`1` 表示插到第 1 页之后
- 空 `page_order` 表示「保持原顺序」（不是「没有页面」）

</details>

<details>
<summary><b>响应示例</b></summary>

```json
{
  "task_id": "09cb844b79d6",
  "filename": "photo_compressed.png",
  "download_url": "/api/download/09cb844b79d6",
  "original_size": 7664509,
  "output_size": 39360,
  "output_format": "png",
  "saved_percent": 99.5,
  "note": ""
}
```

`note` 为空表示无额外说明；转换走兜底通道时会给出说明文字（例如「Word 不可用，已用内置排版引擎生成 PDF」），
前端会直接显示在结果卡片里。

批量接口额外返回 `total_files`、`success_files`、`failed_files`。

</details>

<details>
<summary><b>调用示例</b></summary>

```powershell
# 单文件转换
curl.exe -F "file=@test.docx" -F "target=pdf" http://127.0.0.1:8000/api/convert

# 压缩
curl.exe -F "file=@big.pdf" -F "level=40" http://127.0.0.1:8000/api/compress

# PDF 合并（把第 2 个文件插到主 PDF 第 1 页之后）
curl.exe -F "files=@main.pdf" -F "files=@extra.pdf" -F "mode=insert" -F "insert_index=1" `
  http://127.0.0.1:8000/api/pdf-merge
```

</details>

---

## 项目结构

```
office-converter/
├── main.py              # FastAPI 后端：路由、上传校验、任务与下载
├── converter.py         # 核心转换/压缩/PDF 合并编辑/扫描增强
├── translator.py        # 文献翻译：段落抽取 + 大模型翻译 + Word/PDF 生成
├── office_worker.py     # Office COM 隔离进程入口（超时可控、可清理）
├── docx_pdf.py          # 内置排版引擎（无 Office 时的 docx → PDF 兜底）
├── launcher.py          # start.bat 调用的启动器（提示信息 + 日志）
├── start.bat            # 一键启动（目录名需为 .venv-ocr）
├── glossary.txt         # 默认术语表
├── static/
│   ├── index.html       # 界面结构
│   ├── app.css          # 设计令牌 + 全部样式（含深色主题）
│   └── app.js           # 交互逻辑（含 PDF 可视化编辑器）
├── tests/               # 自检脚本（前端静态检查、JS 语法、PDF 页码语义）
│   └── README.md        # 每个脚本查什么、什么时候该跑
├── docs/
│   └── GITHUB_PUBLISH.md  # 首次发布到 GitHub 的步骤说明
├── .env.example         # 配置模板
├── requirements.txt
├── CHANGELOG.md
├── LICENSE
├── uploads/             # 临时上传（用完即删，已 gitignore）
└── outputs/             # 转换结果（已 gitignore）
```

## 自检

改动前端或 PDF 编辑逻辑后，可以不用浏览器先跑一遍检查：

```powershell
python tests/check_frontend.py          # 括号配平 / id 引用 / 资源路径 / 标签闭合
python tests/check_js_syntax.py         # 用 node 校验 app.js 语法（需要 node）
python tests/test_pdf_page_semantics.py # PDF 编辑与合并的页码语义（需要 PyMuPDF）
python tests/smoke_test.py              # 端到端冒烟：临时起服务跑一遍主要接口
```

`/api/pdf-edit` 的页码约定写错了**不会报错**，只会静默丢页或把标注写到别的页上，
所以要改 `static/app.js` 的 `submitEditor()` 时请先看 [tests/README.md](tests/README.md)。

---

## 已知限制

- **仅支持 Windows**：Office 转 PDF 与 OCR 转 Word 依赖本机 Microsoft Office 的 COM 自动化
  （`pywin32`，已在 `requirements.txt` 中按平台限定）。其余功能理论上跨平台，但未做验证。
- **Office COM 不支持并发**：所有转换任务在单线程池中串行执行，大文件需要排队等待。
- **PDF 压缩是有损的**：中度/强力压缩会把页面重新渲染为 JPEG，**文字会变成图片、不可再搜索**。
  需要保留文字层时请用「轻度」或不要压缩。
- **内置排版引擎的版式差异**：兜底生成的 PDF 保留正文、标题、列表、表格、图片与分页，
  但字体、行距、浮动图片位置会与 Word 有出入。
- **扫描版 PDF 没有文字层**：无法直接改写文字，请先「格式转换 → Word（可编辑）」做 OCR。
- **无鉴权、无 HTTPS**：设计用途是本机或可信局域网自用。

---

## 常见问题

<details>
<summary>转换 Word 时一直转圈或报「Word 不可用」</summary>

Office 自动化有时会被更新、激活弹窗或残留进程卡死。服务已经带了超时保护，不会永久卡住。

1. 在任务管理器结束残留的 `WINWORD.EXE` / `EXCEL.EXE` / `POWERPNT.EXE`
2. 稍等片刻（默认 10 分钟后会自动重新探测），**不需要重启服务**
3. 急用的话可以直接重启服务，或先用「转为可编辑 Word」等其它通道

</details>

<details>
<summary>PDF 编辑里双击文字没反应</summary>

说明这份 PDF 没有文字层（扫描件/图片型 PDF）。请改用编辑面板里的「📝 转为可编辑 Word」先做 OCR，
再编辑 Word；或者直接用工具栏的「文字」「白底遮盖」在 PDF 上叠字。

</details>

<details>
<summary>文献翻译提示未配置密钥</summary>

在项目根目录创建 `.env`（可从 `.env.example` 复制），填入 `TRANSLATE_API_KEY` 后重启服务。
界面右上角的状态徽标会显示「翻译未配置」还是「翻译已配置」。

</details>

<details>
<summary>超过 50 页的 PDF 在编辑器里只看到前 50 页</summary>

预览为了性能只加载前 50 页。**保存时其余页面会原样保留在结果文件末尾**，不会丢失。
前 50 页可以正常排序、删除、旋转和标注。

</details>

---

## 更新日志

见 [CHANGELOG.md](CHANGELOG.md)。

## 许可证

[MIT](LICENSE) © 2026 office-converter contributors
