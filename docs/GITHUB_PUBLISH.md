# 首次发布到 GitHub 操作指南

本文面向**第一次**把 `office-converter` 推到 GitHub 的场景，命令全部为 PowerShell，
可以直接复制粘贴。文中假定项目目录为 `D:\123456\office-converter`，请按自己的实际路径替换。

---

## 0. 发布前先确认三件事

1. **项目是 Windows 专用的。** Word / Excel / PowerPoint 转 PDF 走本机 Microsoft Office 的
   COM 自动化（`pywin32`，`office_worker.py`、`converter.py` 里的 `_run_office_isolated`），
   没有 Office 的 Linux/macOS 机器上这部分功能不可用。请在 README 顶部和下面的「仓库描述」
   里明确写出 **Windows-only**，避免别人 clone 下来发现跑不通。
2. **不要提交 `.env`。** 根目录的 `.env` 里放的是真实 `TRANSLATE_API_KEY`。
   `.gitignore` 已经排除它，仓库里只保留 `.env.example` 模板。
3. **确认当前还没有 git 仓库。** 本目录目前没有 `.git`，属于首次初始化。

---

## 1. 哪些文件该提交，哪些不该提交

### 应该提交（源码与文档）

| 路径 | 说明 |
|------|------|
| `main.py` | FastAPI 后端，所有 API 端点 |
| `converter.py` | 核心转换 / 压缩 / PDF 工具 / 扫描增强 |
| `translator.py` | 英文文献翻译（段落抽取 + 大模型翻译 + Word 生成） |
| `docx_pdf.py` | 不依赖 Office 的 docx → PDF 兜底排版引擎 |
| `office_worker.py` | Office COM 隔离子进程入口 |
| `launcher.py` | 启动入口（中文提示、日志写入 `_start_log.txt`） |
| `start.bat` | 一键启动（保持纯 ASCII，中文提示都在 launcher.py 里） |
| `static/index.html` | 前端界面结构 |
| `static/app.css` | 前端样式（设计令牌 + 深色主题） |
| `static/app.js` | 前端交互逻辑（含 PDF 可视化编辑器） |
| `tests/` | 无需浏览器的自检脚本（见 `tests/README.md`） |
| `glossary.txt` | 默认术语表（翻译功能自动加载） |
| `requirements.txt` | 依赖清单 |
| `.env.example` | 环境变量模板（**必须提交**，里面不能有真实密钥） |
| `.gitignore` / `LICENSE` / `CHANGELOG.md` / `README.md` / `docs/` | 项目文档与协议 |

### 绝对不能提交

| 路径 | 原因 |
|------|------|
| `.venv-ocr/` | 本地虚拟环境，**上万个文件**，且含平台相关的 `.exe`/`.pyd` 与绝对路径信息。别人用 `pip install -r requirements.txt` 自己建即可，提交它既臃肿又会让 clone 变慢 |
| `outputs/` | **转换结果目录**，本次检查时该目录下已累积上万个产物文件（整页 JPG、PDF、批量 zip），体积巨大且多为本地业务数据/隐私文档。发布前务必确认它被忽略 |
| `uploads/` | 临时上传目录，同样含真实文档内容 |
| `.env` | 真实 API 密钥 |
| `_start_log.txt`、`*.log` | 运行日志 |
| `__pycache__/`、`*.pyc` | Python 字节码缓存 |
| `_probe_*`、`_t_*`、`_selflaunch*`、`_medium_test.bat` 等 | 开发期探测/测试脚本残留 |
| `large-upload-response.json`、`real-preview-response.json` | 调试时保存的接口响应 |
| `原版式翻译预览_p*.png` | 调试用预览图 |
| `revise_draft.py`、`source_docs/` | 一次性处理本机手稿的脚本，与程序运行无关 |
| `.vscode/`、`.idea/`、`.DS_Store`、`Thumbs.db` | 编辑器 / 系统垃圾文件 |

以上内容都已在 `.gitignore` 中列出。

### 发布前自查（重要）

在第一次 `git add` 之前，先看一眼到底会提交多少文件：

```powershell
cd D:\123456\office-converter
git status --short | Measure-Object -Line
```

如果这个数字是几千甚至上万，说明 `.gitignore` 没生效（常见原因是文件不在仓库根目录，
或者条目被写成了 `outputs` 却没有生效于目录）——**先解决这个问题再继续**，
否则会把海量结果文件写进历史，之后清理 history 非常麻烦。

也可以确认某个大目录确实被忽略（被忽略的路径不会出现在输出里）：

```powershell
git check-ignore -v outputs .venv-ocr uploads .env
```

---

## 2. 首次发布：完整命令（PowerShell）

```powershell
# 进入项目目录
cd D:\123456\office-converter

# 1) 初始化仓库，主分支命名为 main
#    （git 2.28 以下不支持 -b，先执行 git init，再执行 git branch -M main）
git init -b main

# 2) 确认身份已配置（没配过才需要执行这两行，邮箱/名字换成自己的）
git config user.name  "your-name"
git config user.email "your-email@example.com"

# 3) 把该提交的文件加入暂存区
git add .

# 4) 检查将要提交的文件清单：重点确认没有 outputs/、.venv-ocr/、.env
git status --short

# 5) 提交
git commit -m "feat: 办公文档转换器 1.0.0 首个公开版本"

# 6) 先在 GitHub 网页上新建一个空仓库（不要勾选 Add README / .gitignore / license），
#    假设仓库地址是 https://github.com/<你的用户名>/office-converter.git
git remote add origin https://github.com/<你的用户名>/office-converter.git

# 7) 推送
git push -u origin main
```

推送时如果弹出凭据窗口：用户名填 GitHub 账号，密码位置要填 **Personal Access Token**
（GitHub 从 2021 年起不再接受账号密码）。生成路径：
GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens，
权限勾选该仓库的 **Contents: Read and write**。

如果本机装了 GitHub CLI，可以省掉网页建仓库这一步：

```powershell
gh auth login
gh repo create office-converter --public --source . --remote origin --push
```

---

## 3. 打 tag（可选，与 CHANGELOG 对应）

```powershell
cd D:\123456\office-converter
git tag -a v1.0.0 -m "v1.0.0"
git push origin v1.0.0
```

推完 tag 后可以在 GitHub 的 Releases 页面基于这个 tag 发布版本说明。
`CHANGELOG.md` 里的 `[1.0.0]` 目前日期是 `2026-XX-XX` 占位，正式发布前请改成真实日期。

---

## 4. 仓库描述（Description）与 Topics 建议

### About / Description（一句话，建议 120 字符内）

推荐：

> Windows 本地办公文档转换器：Word/Excel/PPT/PDF/图片互转，支持压缩、PDF 合并与可视化编辑、扫描增强和英文文献大模型翻译。

再准备一条英文版（GitHub 上国际访问者更多，中英都写更实用）：

> Windows-only office document converter: Word/Excel/PPT/PDF/image conversion, compression, PDF merge & visual editing, scan enhancement, and LLM-based EN→ZH literature translation.

### Topics（标签，最多 20 个，按需取用）

```
fastapi
python
document-converter
pdf
docx
xlsx
pptx
pdf-editing
pdf-merge
image-conversion
ocr
tesseract
translation
llm
deepseek
office-automation
com-automation
windows-only
self-hosted
vanilla-javascript
```

选择原则：前 5 个放最能说明「这是什么 + 用什么写的」的标签（`fastapi`、`python`、
`document-converter`、`pdf`、`windows-only`）；`windows-only` 和 `self-hosted` 一定要有，
它们直接告诉别人运行环境与用途。

### README 顶部建议补充的三行

GitHub 仓库页第一屏决定别人是否继续看，建议在 README 最前面加：

```markdown
# 办公文档转换器

Windows 本地运行的办公文档转换服务（FastAPI + 原生 JS），浏览器打开 http://127.0.0.1:8000 使用。

> ⚠️ 仅支持 Windows：Office 文档转 PDF 依赖本机安装的 Microsoft Office（COM 自动化）。
> 未安装 Office 时，图片 / PDF / 文本类功能仍可正常使用。
```

---

## 5. 常见问题

**Q：`.gitignore` 写好了，但 `git status` 里还是能看到 `outputs/` 怎么办？**
说明该目录已经被加进暂存区了，先撤出来：

```powershell
git rm -r --cached outputs
git rm -r --cached .venv-ocr
git commit -m "chore: 从版本控制中移除结果目录与虚拟环境"
```

如果这些文件已经 push 上去了，只删文件是不够的（历史里还在），需要 `git filter-repo`
或 BFG 清理历史；密钥一旦被 push，请立刻去服务商后台**吊销并重新生成**。

**Q：clone 下来的人怎么运行？**

```powershell
git clone https://github.com/<你的用户名>/office-converter.git
cd office-converter
python -m venv .venv-ocr
.\.venv-ocr\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env   # 需要文献翻译时才填 TRANSLATE_API_KEY
start.bat
```

**Q：文献翻译提示「尚未配置翻译密钥」？**
`.env` 优先由 `python-dotenv` 读取（已在 `requirements.txt` 中）；
即使没有安装该包，`main.py` 也会用内置的极简解析器兜底，所以下面两种方式都可行：

```powershell
# 方式一：写进 .env
Copy-Item .env.example .env
# 编辑 .env，填入 TRANSLATE_API_KEY=sk-xxxx，然后重启服务

# 方式二：直接用环境变量
$env:TRANSLATE_API_KEY = "sk-xxxx"
python main.py
```

**Q：API 密钥会不小心被提交吗？**
`.gitignore` 里已包含 `.env`。养成习惯：每次 `git add .` 之后、`git commit` 之前
都跑一次 `git status --short` 快速扫一眼文件清单。
