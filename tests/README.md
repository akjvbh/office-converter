# 自检脚本

这些脚本不需要浏览器、不需要启动服务，用来在改动前端或 PDF 逻辑后快速自查。
全部只用标准库，**不需要 pytest**，退出码非 0 表示有检查失败。

```powershell
# 1. 前端静态检查：括号配平、id 引用、资源路径、标签闭合（零依赖）
python tests/check_frontend.py

# 2. 用 node 校验 static/app.js 的语法（需要 node；没有则跳过，退出码 2）
python tests/check_js_syntax.py

# 3. PDF 编辑 / 合并的页码语义（需要 PyMuPDF，即 requirements.txt 里的 PyMuPDF）
python tests/test_pdf_page_semantics.py

# 4. 真实端到端冒烟测试：临时在 8123 端口起服务并跑一遍主要接口（需要 fastapi/uvicorn）
python tests/smoke_test.py
```

第 4 个脚本会自行启动和关闭服务，不会影响你正在运行的服务（用的是 8123 端口），
覆盖：首页与静态资源、`/api/config`、`/api/formats`、压缩、PDF 删除页、PDF 重排+标注、
PDF 预览、PDF 合并、批量压缩，以及错误路径（不支持的目标格式 400、不存在的任务 404）。

## 为什么有第 3 个脚本

`/api/pdf-edit` 的页码约定很容易写错，而且错了以后**不会报错**，只会静默丢页或把标注
写到别的页上。这个脚本把这些约定固定下来：

- `page_order` 用的是**原始 1-based 页码**，而且同时是白名单（没列出的页会被丢弃）
- `delete_pages` / `rotations` 也用原始 1-based 页码，`rotations` 的值是**增量角度**
- `annotations` 的 `page` 是 **0-based**，指的是「该页在 `page_order` 中的位置」
  —— 因为后端先写标注、再做重排
- 超过 50 页时，前端只预览前 50 页，保存时要把未预览的页面追加回 `page_order`，否则会丢页
- 合并接口的 `insert_index` 是 **0-based**

改动 `static/app.js` 的 `submitEditor()` / `currentOrder()`，或改动 `converter.edit_pdf()`
之后，请务必跑一次这个脚本。
