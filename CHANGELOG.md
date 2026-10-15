# 更新日志

本文件记录本项目的所有重要变更。

格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Changed

- **前端拆分为三个文件**：`static/index.html`（结构）、`static/app.css`（设计令牌 + 全部样式）、
  `static/app.js`（交互逻辑）。原先 1199 行的单文件就地拆开，样式与逻辑不再混在一起。
- **界面重做**：新增深色/浅色主题（跟随系统并可手动切换）、统一的状态与错误展示区、
  批量失败清单与错误详情可一键复制、结果卡片显示更多元数据。
- **切换功能不再清空已选文件**：在选项卡之间切换时会保留当前文件，可直接换一种方式处理同一批文件。
- **PDF 编辑器改为整篇可撤销**：标注、删页、旋转、页面重排共用同一个撤销栈，
  不再出现「撤销只回退了部分操作」的情况。
- **压缩预估不再与最终结果脱节**：选文件后不再自动跑 4 次压缩；
  改为点击「预估四档压缩后的大小」后展示四档结果，且可直接下载其中任意一档（该档压缩只执行一次）。

### Fixed

- **PDF 编辑器删除页后标注错位**：删除页面时同步重排标注的页序，避免标注落到错误的页上。
- **重排页面后标注落错页**：提交时把标注的页序换算成「该页在 `page_order` 中的位置」，
  修正了复制页/重排后标注写到别的页面上的问题。
- **超过 50 页的 PDF 保存后会丢掉其余页面**：现在只加载前 50 页用于预览，
  保存时未预览的页面会按原顺序追加回结果文件。
- **页面旋转后无法在该页添加标注**：旋转改为在画布上真正重绘，
  点击坐标与绘制坐标都会在「PDF 坐标」与「旋转后显示坐标」之间换算。
- **没有标注时也会整篇重写 PDF**：后端只在 `annotations` 非空时才走
  「先擦除再写入」的通道，避免空标注触发全页 `apply_redactions`。
- **`.env` 未安装 `python-dotenv` 时静默失效**：`main.py` 新增内置的极简 `.env` 解析器兜底，
  并把 `python-dotenv` 加入 `requirements.txt`。

## [1.0.0] - 2026-XX-XX

首个公开版本：Windows 上的本地办公文档转换服务，FastAPI 后端 + 原生 JS 前端，
浏览器打开 `http://127.0.0.1:8000` 即可使用。

### Added

- **格式转换**（`POST /api/convert`，映射见 `GET /api/formats`）：
  - Word：`doc` → pdf/docx/txt；`docx` → pdf/txt/html
  - Excel：`xls` → pdf/xlsx/csv；`xlsx` → pdf/csv
  - PowerPoint：`ppt` → pdf/pptx；`pptx` → pdf/png（多页自动打包 zip）
  - PDF：→ docx/png/jpg/txt
  - CSV → xlsx；Markdown → html/pdf；TXT → pdf/docx
  - 图片（png/jpg/jpeg/jfif/bmp/webp/gif/tiff）互转，并可转 PDF、ICO
- **图片转可编辑 Word（OCR）**：图片 → `docx`。用 Tesseract（`pytesseract`，
  优先 `chi_sim+eng`，缺中文语言包时回退 `eng`）识别文字与坐标，再由 Word COM 生成
  「扫描图作底图 + 文字浮动文本框」的可编辑文档；该功能需要本机安装 Tesseract 和 Word。
- **Office COM 转 PDF**：Word/Excel/PowerPoint 转 PDF 走本机 Microsoft Office COM 自动化，
  通过独立子进程 `office_worker.py` 执行（操作类型 `pdf` / `probe` / `format` / `scan-docx`），
  并带健康探测、超时与残留进程清理。
- **内置排版引擎兜底**：Word 不可用时，`docx` → PDF 自动改用 `docx_pdf.py`
  （python-docx 解析 + PyMuPDF 排版）生成 PDF，结果中通过 `note` 字段提示「已用内置排版引擎」。
- **文件压缩**（`POST /api/compress`）：四档级别 `80/60/40/20`；图片做有损压缩并限制最长边
  （4096/2560/1600/1200 px，PNG 另做调色板量化），PDF 把每页按 150/110/80/60 dpi 重渲染为
  JPEG（质量 80/60/40/20），其它类型打包 zip；压缩后反而变大时会回退为原文件。
- **批量处理**（`POST /api/batch`）：`mode=convert|compress`，多文件重名自动加序号，
  单个结果直接返回、多个结果打包 zip，并返回成功/失败文件清单。
- **PDF 合并**（`POST /api/pdf-merge`）：支持 `sequence`（顺序）、`alternate`（逐页交替）、
  `prepend`（前插）、`append`（后插）、`insert`（插入指定页后，`insert_index` 从 0 开始）。
- **PDF 页面编辑**（`POST /api/pdf-edit`）：删除页、旋转页（`{"1":90,"3":180}`）、页面重排，
  以及 `annotations` 标注写入，支持 `replace_text`（先 redaction 移除原文字再写入新文字）、
  `text`（加字）、`rect` / `ellipse` / `line` / `arrow`、`highlight` / `underline` / `strike`、
  `image`（插入图片）和 `signature`（自由绘制签名）。
- **PDF 可视化预览**（`POST /api/pdf-preview`）：返回页面 PNG（base64）与可编辑文字区域
  （文字、坐标矩形、字号），最多预览前 50 页，超过时返回 `truncated`。
- **一键 PDF 文件夹转 JPG**（`POST /api/pdf-folder-to-jpg`）：批量接收 PDF，
  逐页按 144 dpi 渲染为 JPG，按 82→5 逐级降低 JPEG 质量、仍超限则按 0.8 倍逐级缩小尺寸，
  保证每张不超过 300KB，最后打包 zip；响应含 PDF 数、图片数与失败文件清单。
- **扫描照片增强**（`POST /api/scan-photo`）：Canny 边缘 + 四边形轮廓自动定位纸张四角并做
  透视矫正（未找到四边形时退回原图），随后用大核高斯背景做光照归一化、双边滤波降噪，
  输出灰度扫描 JPG，响应含 `document_detected`、`width`、`height`。
- **英文文献翻译**（`POST /api/translate` + `GET /api/translate/progress/{task_id}`）：
  从 PDF/DOCX/TXT 抽取段落，逐段调用 OpenAI 兼容的 `/chat/completions` 接口翻译，
  支持 `bilingual`（中英对照）/`chinese`（仅译文）/`both`（两者打包 zip）/
  `layout_pdf`（原版式 PDF，擦除原文并在原坐标写入中文）四种输出。
  可传 `domain` 专业领域、`glossary` 术语表、`translate_names` 人名是否译成中文、
  `skip_references` 是否跳过参考文献；任务异步执行并通过进度接口轮询。
- **默认术语表**：根目录 `glossary.txt` 每行 `英文=中文`（`#` 为注释），
  翻译时后端自动携带，前端通过 `GET /api/config` 预填输入框，修改后无需重启即生效。
- **公开配置接口**（`GET /api/config`）：返回 `max_file_size_mb`、`max_batch_size_mb`、
  `translate_configured`、`translate_model`、`default_glossary`，供前端展示与大小校验。
- **结果下载**（`GET /api/download/{task_id}`）：转换/压缩结果按任务 ID 下载，
  响应包含原始大小、输出大小、输出格式、体积减小百分比与可选 `note` 提示。
- **一键启动脚本**：`start.bat`（纯 ASCII，避免 cmd 解析问题）调用 `launcher.py`，
  使用 `.venv-ocr` 解释器启动服务，并把输出同时打印到窗口和写入 `_start_log.txt`。
- **前端界面**：`static/index.html` 单页原生 JS 界面，支持拖拽/多选/整文件夹选择、
  转换与压缩选项卡、PDF 工具、扫描照片、文献翻译进度条。
  （1.0.0 时为单文件实现，后续已拆分为 `index.html` + `app.css` + `app.js`，见 Unreleased。）

### Configuration

- 环境变量（可用项目根目录 `.env` 提供，模板见 `.env.example`）：
  - `MAX_FILE_SIZE_MB`（默认 200）、`MAX_BATCH_SIZE_MB`（默认 500），必须是大于 0 的整数
  - `OFFICE_TIMEOUT_SECONDS`（默认 900）、`OFFICE_PROBE_TIMEOUT_SECONDS`（默认 45）、
    `OFFICE_RECHECK_SECONDS`（默认 600）
  - `TRANSLATE_API_KEY`、`TRANSLATE_BASE_URL`（默认 `https://api.deepseek.com/v1`）、
    `TRANSLATE_MODEL`（默认 `deepseek-chat`），以及可选的 `TRANSLATE_TARGET_LANG`、
    `TRANSLATE_TEMPERATURE`、`TRANSLATE_MAX_CHARS`、`TRANSLATE_TIMEOUT`、
    `TRANSLATE_SKIP_REFERENCES`、`TRANSLATE_WORKERS`
- 未配置 `TRANSLATE_API_KEY` 时只有文献翻译不可用，其它功能不受影响。
- `.env` 优先由 `python-dotenv` 读取；没有安装该包时 `main.py` 会用内置的极简解析器兜底，
  因此「复制 `.env.example` 为 `.env` 填密钥」不依赖额外安装。

### Known limitations

- 仅支持 Windows：Office 转 PDF 依赖本机 Microsoft Office 的 COM 自动化
  （`pywin32`，`requirements.txt` 中已按 `sys_platform == 'win32'` 限定）。
- Office COM 不支持并发，所有转换任务在单线程池中串行执行，大文件需要等待。
- 内置排版引擎无法解析旧版二进制 `.doc`，遇到 `.doc` 需先用 Word 另存为 `.docx`。
- OCR 图片转可编辑 Word 同样依赖 Word COM，没有兜底通道：Word 不可用时会直接报错提示。
- 「原版式 PDF」翻译只支持 PDF 输入；扫描版 PDF 没有文字层，需先转为可编辑 Word。
- 无鉴权、无 HTTPS，设计用途是本机或局域网内自用。

[Unreleased]: https://github.com/OWNER/office-converter/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/OWNER/office-converter/releases/tag/v1.0.0
