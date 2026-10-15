"""
办公文档转换器 - FastAPI 后端
运行: uvicorn main:app --host 0.0.0.0 --port 8000
或:   python main.py
"""
import os
import uuid
import shutil
import asyncio
import base64
import json
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

import converter
import translator

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
STATIC_DIR = os.path.join(BASE_DIR, "static")
GLOSSARY_FILE = os.path.join(BASE_DIR, "glossary.txt")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _default_glossary() -> str:
    """项目根目录 glossary.txt 里的默认术语表；每次读取，改完即时生效（无需重启）。"""
    try:
        with open(GLOSSARY_FILE, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError:
        return ""


def _merged_glossary(user_text: str) -> str:
    """默认术语表在前、界面填写的在后（界面内容可以覆盖或补充）。"""
    default = _default_glossary()
    user_text = (user_text or "").strip()
    if default and user_text:
        return default + "\n" + user_text
    return default or user_text


def _load_env_file(path: str) -> None:
    """读取 .env 并写入 os.environ（已存在的真实环境变量优先）。

    优先用 python-dotenv；没装时用内置的极简解析器兜底，
    这样「复制 .env.example 为 .env」的用法不依赖额外安装。
    """
    if not os.path.isfile(path):
        return
    try:
        from dotenv import load_dotenv
        load_dotenv(path)
        return
    except ImportError:
        pass
    try:
        with open(path, "r", encoding="utf-8") as handle:
            for raw_line in handle:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, value = line.partition("=")
                key = key.strip()
                if key.lower().startswith("export "):
                    key = key[7:].strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except OSError as exc:
        print(f"[env] 无法读取 {path}: {exc}")


# 读取项目根目录的 .env（若存在）以配置翻译密钥；已存在的真实环境变量优先
_load_env_file(os.path.join(BASE_DIR, ".env"))

# 文献翻译配置：未配置密钥时仅该功能不可用，不影响其它转换
TRANSLATE_CFG = translator.config_from_env()
TRANSLATE_SETUP_HINT = (
    "尚未配置翻译密钥。请在项目根目录创建 .env 文件并填写：\n"
    "TRANSLATE_API_KEY=sk-xxxx\n"
    "TRANSLATE_BASE_URL=https://api.deepseek.com/v1\n"
    "TRANSLATE_MODEL=deepseek-chat\n"
    "保存后重启服务即可。"
)


def _positive_int_env(name: str, default: int) -> int:
    """读取正整数环境变量，配置错误时在启动阶段给出明确提示。"""
    raw = os.getenv(name, str(default))
    try:
        value = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"环境变量 {name} 必须是正整数，当前值: {raw!r}") from exc
    if value <= 0:
        raise RuntimeError(f"环境变量 {name} 必须大于 0，当前值: {value}")
    return value


MAX_FILE_SIZE_MB = _positive_int_env("MAX_FILE_SIZE_MB", 200)
MAX_BATCH_SIZE_MB = _positive_int_env("MAX_BATCH_SIZE_MB", 500)
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024
MAX_BATCH_SIZE = MAX_BATCH_SIZE_MB * 1024 * 1024

app = FastAPI(title="办公文档转换器")
# Office COM 不支持多线程并发，用单线程池串行执行转换任务
executor = ThreadPoolExecutor(max_workers=1)
# 文献翻译以网络请求为主，用独立线程池，避免长时间占用 Office 的串行队列
translate_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="translate")

# 任务结果记录: task_id -> 输出文件路径
_results: dict[str, str] = {}
# 翻译任务进度: task_id -> {state, done, total, message, result}
_translate_tasks: dict[str, dict] = {}
# 持有后台任务引用，防止被垃圾回收
_translate_jobs: dict[str, object] = {}


async def _save_pdf_files(files: list[UploadFile], task_id: str):
    """保存 PDF 工具上传的多个文件，并执行与普通批量上传相同的大小限制。"""
    upload_dir = os.path.join(UPLOAD_DIR, task_id)
    output_dir = os.path.join(OUTPUT_DIR, task_id)
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    paths, total = [], 0
    try:
        for number, file in enumerate(files):
            name = os.path.basename(file.filename or f"document_{number + 1}.pdf")
            if not name.lower().endswith(".pdf"):
                raise HTTPException(400, f"仅支持 PDF 文件：{name}")
            path = os.path.join(upload_dir, f"{number}_{name}")
            size = 0
            with open(path, "wb") as stream:
                while chunk := await file.read(1024 * 1024):
                    size += len(chunk); total += len(chunk)
                    if size > MAX_FILE_SIZE or total > MAX_BATCH_SIZE:
                        raise HTTPException(413, "PDF 文件超过上传大小限制")
                    stream.write(chunk)
            paths.append(path)
    except Exception:
        shutil.rmtree(upload_dir, ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)
        raise
    return paths, upload_dir, output_dir, total


async def _run_pdf_operation(files, operation, **kwargs):
    task_id = uuid.uuid4().hex[:12]
    paths, upload_dir, output_dir, total = await _save_pdf_files(files, task_id)
    out_path = os.path.join(output_dir, "result.pdf")
    loop = asyncio.get_event_loop()
    try:
        await loop.run_in_executor(executor, lambda: operation(paths, out_path, **kwargs))
    except Exception as exc:
        shutil.rmtree(output_dir, ignore_errors=True)
        raise HTTPException(500, f"PDF 处理失败: {exc}")
    finally:
        shutil.rmtree(upload_dir, ignore_errors=True)
    _results[task_id] = out_path
    return _build_result(task_id, out_path, total)


def _merge_operation(paths, out_path, mode="sequence", insert_index=0):
    return converter.merge_pdfs(paths, out_path, mode, insert_index)


def _edit_operation(paths, out_path, delete_pages=None, rotations=None, page_order=None):
    if len(paths) != 1:
        raise ValueError("编辑模式只能上传一个 PDF")
    return converter.edit_pdf(paths[0], out_path, delete_pages, rotations, page_order)


def _annotate_operation(paths, out_path, annotations):
    return converter.annotate_pdf(paths[0], out_path, annotations)


def _edit_and_annotate_operation(paths, out_path, delete_pages=None, rotations=None,
                                  page_order=None, annotations=None):
    """先写入文字编辑/标注，再执行页面删除、旋转和重排，避免参数互相覆盖。"""
    if len(paths) != 1:
        raise ValueError("编辑模式只能上传一个 PDF")
    temp_path = out_path + ".annotated.pdf"
    try:
        converter.annotate_pdf(paths[0], temp_path, annotations or [])
        converter.edit_pdf(temp_path, out_path, delete_pages, rotations, page_order)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


@app.get("/", response_class=HTMLResponse)
async def index():
    with open(os.path.join(STATIC_DIR, "index.html"), "r", encoding="utf-8") as f:
        return f.read()


@app.get("/api/formats")
async def formats():
    """返回所有支持的转换映射"""
    return converter.CONVERSIONS


@app.get("/api/config")
async def public_config():
    """返回前端需要展示和校验的公开配置。"""
    return {
        "max_file_size_mb": MAX_FILE_SIZE_MB,
        "max_batch_size_mb": MAX_BATCH_SIZE_MB,
        "translate_configured": TRANSLATE_CFG.configured,
        "translate_model": TRANSLATE_CFG.model,
        # 供界面预填「术语表」；实际翻译时后端也会自动带上（见 _merged_glossary）
        "default_glossary": _default_glossary(),
    }


@app.post("/api/pdf-merge")
async def pdf_merge(
    files: list[UploadFile] = File(...),
    mode: str = Form("sequence"),
    insert_index: int = Form(0),
):
    """合并多个 PDF：顺序、逐页交替、前插、后插或指定页码插入。"""
    if mode not in {"sequence", "alternate", "prepend", "append", "insert"}:
        raise HTTPException(400, "无效的 PDF 合并模式")
    if len(files) < 2:
        raise HTTPException(400, "至少选择两个 PDF")
    return await _run_pdf_operation(files, _merge_operation, mode=mode, insert_index=insert_index)


@app.post("/api/pdf-edit")
async def pdf_edit(
    file: UploadFile = File(...),
    delete_pages: str = Form(""),
    rotations: str = Form("{}"),
    page_order: str = Form(""),
    annotations: str = Form("[]"),
):
    """编辑 PDF 页面：删除、旋转及重排。页码均从 1 开始。"""
    import json
    try:
        deleted = [int(x.strip()) for x in delete_pages.split(",") if x.strip()]
        rotate_map = json.loads(rotations or "{}")
        order = [int(x.strip()) for x in page_order.split(",") if x.strip()] or None
        marks = json.loads(annotations or "[]")
        if not isinstance(marks, list):
            raise ValueError("annotations 必须是数组")
    except (ValueError, json.JSONDecodeError) as exc:
        raise HTTPException(400, "页码参数格式错误") from exc
    if marks:
        return await _run_pdf_operation(
            [file], _edit_and_annotate_operation, delete_pages=deleted,
            rotations=rotate_map, page_order=order, annotations=marks
        )
    return await _run_pdf_operation([file], _edit_operation, delete_pages=deleted,
                                    rotations=rotate_map, page_order=order)


@app.post("/api/pdf-preview")
async def pdf_preview(file: UploadFile = File(...)):
    """返回 PDF 页面预览图，供可视化编辑器叠加操作。"""
    task_id, src_path, task_output, src_size = await _save_upload(file)
    try:
        import fitz
        doc = fitz.open(src_path)
        pages = []
        page_count = len(doc)
        for page in doc[:50]:
            pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
            text_items = []
            page_dict = page.get_text("dict")
            for block in page_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        if text.strip():
                            text_items.append({
                                "text": text,
                                "rect": [round(v, 2) for v in span["bbox"]],
                                "size": round(float(span.get("size", 11)), 1),
                            })
            pages.append({"width": page.rect.width, "height": page.rect.height,
                          "image": "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode(),
                          "texts": text_items})
        doc.close()
        return {"pages": pages, "page_count": page_count, "truncated": page_count > 50}
    except Exception as exc:
        raise HTTPException(500, f"PDF 预览失败: {exc}")
    finally:
        shutil.rmtree(os.path.dirname(src_path), ignore_errors=True)


async def _save_upload(file: UploadFile) -> tuple[str, str, str, int]:
    """保存上传文件，返回 (task_id, src_path, task_output, size)"""
    filename = os.path.basename(file.filename or "")
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    if not ext:
        raise HTTPException(400, "文件缺少扩展名")

    task_id = uuid.uuid4().hex[:12]
    task_upload = os.path.join(UPLOAD_DIR, task_id)
    task_output = os.path.join(OUTPUT_DIR, task_id)
    os.makedirs(task_upload, exist_ok=True)
    os.makedirs(task_output, exist_ok=True)

    src_path = os.path.join(task_upload, filename)
    size = 0
    with open(src_path, "wb") as f:
        while chunk := await file.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_FILE_SIZE:
                f.close()
                shutil.rmtree(task_upload, ignore_errors=True)
                shutil.rmtree(task_output, ignore_errors=True)
                raise HTTPException(413, f"文件超过 {MAX_FILE_SIZE_MB}MB 限制")
            f.write(chunk)
    return task_id, src_path, task_output, size


def _build_result(task_id: str, out_path: str, src_size: int) -> dict:
    """构造带大小/格式标签的响应"""
    out_size = os.path.getsize(out_path)
    ratio = round((1 - out_size / src_size) * 100, 1) if src_size else 0
    return {
        "task_id": task_id,
        "filename": os.path.basename(out_path),
        "download_url": f"/api/download/{task_id}",
        "original_size": src_size,
        "output_size": out_size,
        "output_format": os.path.splitext(out_path)[1].lstrip(".").lower(),
        "saved_percent": ratio,
        # 转换方式说明（例如 Word 不可用、改用内置排版引擎生成 PDF）
        "note": converter.pop_engine_note(out_path),
    }


@app.post("/api/convert")
async def convert_file(file: UploadFile = File(...), target: str = Form(...)):
    filename = os.path.basename(file.filename or "")
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    targets = converter.get_supported_targets(ext)
    if target.lower() not in targets:
        raise HTTPException(400, f"不支持 .{ext} 转换为 .{target}，可选: {targets}")

    task_id, src_path, task_output, src_size = await _save_upload(file)
    task_upload = os.path.dirname(src_path)

    loop = asyncio.get_event_loop()
    try:
        out_path = await loop.run_in_executor(
            executor, converter.convert, src_path, task_output, target.lower()
        )
    except Exception as e:
        shutil.rmtree(task_output, ignore_errors=True)
        raise HTTPException(500, f"转换失败: {e}")
    finally:
        # 上传的源文件用完即删
        shutil.rmtree(task_upload, ignore_errors=True)

    _results[task_id] = out_path
    return _build_result(task_id, out_path, src_size)


@app.post("/api/compress")
async def compress_file(file: UploadFile = File(...), level: str = Form("60")):
    if level not in converter.COMPRESS_LEVELS:
        raise HTTPException(400, f"无效压缩级别，可选: {sorted(converter.COMPRESS_LEVELS)}")

    task_id, src_path, task_output, src_size = await _save_upload(file)
    task_upload = os.path.dirname(src_path)

    loop = asyncio.get_event_loop()
    try:
        out_path = await loop.run_in_executor(
            executor, converter.compress_file, src_path, task_output, level
        )
    except Exception as e:
        shutil.rmtree(task_output, ignore_errors=True)
        raise HTTPException(500, f"压缩失败: {e}")
    finally:
        shutil.rmtree(task_upload, ignore_errors=True)

    _results[task_id] = out_path
    return _build_result(task_id, out_path, src_size)



def _batch_process(srcs: list, task_output: str, mode: str, target: str, level: str):
    """批量处理（在线程池中执行），返回 (输出路径列表, 失败文件名列表)"""
    outputs, failed = [], []
    for src in srcs:
        try:
            if mode == "convert":
                ext = os.path.splitext(src)[1].lower().lstrip(".")
                if target.lower() not in converter.get_supported_targets(ext):
                    failed.append(os.path.basename(src) + "（不支持转为 ." + target + "）")
                    continue
                outputs.append(converter.convert(src, task_output, target.lower()))
            else:
                outputs.append(converter.compress_file(src, task_output, level))
        except Exception as e:
            failed.append(os.path.basename(src) + f"（{e}）")
    return outputs, failed


@app.post("/api/batch")
async def batch(
    files: list[UploadFile] = File(...),
    mode: str = Form(...),                 # convert | compress
    target: str = Form(""),
    level: str = Form("60"),
):
    if mode not in ("convert", "compress"):
        raise HTTPException(400, "mode 必须是 convert 或 compress")
    if mode == "convert" and not target:
        raise HTTPException(400, "convert 模式需要 target 参数")
    if mode == "compress" and level not in converter.COMPRESS_LEVELS:
        raise HTTPException(400, "无效压缩级别")

    task_id = uuid.uuid4().hex[:12]
    task_upload = os.path.join(UPLOAD_DIR, task_id)
    task_output = os.path.join(OUTPUT_DIR, task_id)
    work_dir = os.path.join(task_output, "_work")
    os.makedirs(task_upload, exist_ok=True)
    os.makedirs(work_dir, exist_ok=True)

    # 保存所有上传文件（重名自动加序号）
    srcs, total_in = [], 0
    used_names = set()
    for uf in files:
        name = os.path.basename(uf.filename or "")
        if not name or "." not in name:
            continue
        base, ext = os.path.splitext(name)
        i = 1
        while name in used_names:
            name = f"{base}_{i}{ext}"
            i += 1
        used_names.add(name)
        p = os.path.join(task_upload, name)
        file_size = 0
        with open(p, "wb") as f:
            while chunk := await uf.read(1024 * 1024):
                file_size += len(chunk)
                total_in += len(chunk)
                if file_size > MAX_FILE_SIZE:
                    f.close()
                    shutil.rmtree(task_upload, ignore_errors=True)
                    shutil.rmtree(task_output, ignore_errors=True)
                    raise HTTPException(413, f"文件 {name} 超过 {MAX_FILE_SIZE_MB}MB 限制")
                if total_in > MAX_BATCH_SIZE:
                    f.close()
                    shutil.rmtree(task_upload, ignore_errors=True)
                    shutil.rmtree(task_output, ignore_errors=True)
                    raise HTTPException(413, f"批量总大小超过 {MAX_BATCH_SIZE_MB}MB 限制")
                f.write(chunk)
        srcs.append(p)

    if not srcs:
        shutil.rmtree(task_upload, ignore_errors=True)
        shutil.rmtree(task_output, ignore_errors=True)
        raise HTTPException(400, "没有有效文件")

    loop = asyncio.get_event_loop()
    try:
        outputs, failed = await loop.run_in_executor(
            executor, _batch_process, srcs, work_dir, mode, target, level
        )
    finally:
        shutil.rmtree(task_upload, ignore_errors=True)

    if not outputs:
        shutil.rmtree(task_output, ignore_errors=True)
        raise HTTPException(500, "全部文件处理失败: " + "; ".join(failed[:5]))

    # 单个输出直接返回该文件，多个打包 zip
    if len(outputs) == 1:
        out_path = os.path.join(task_output, os.path.basename(outputs[0]))
        shutil.move(outputs[0], out_path)
    else:
        label = "converted" if mode == "convert" else "compressed"
        out_path = os.path.join(task_output, f"batch_{label}_{len(outputs)}files.zip")
        converter.zip_files(outputs, out_path)
    shutil.rmtree(work_dir, ignore_errors=True)

    _results[task_id] = out_path
    result = _build_result(task_id, out_path, total_in)
    result["total_files"] = len(srcs)
    result["success_files"] = len(outputs)
    result["failed_files"] = failed
    return result


def _pdf_folder_to_jpgs(srcs: list, work_dir: str):
    """将所有 PDF 的每一页转为不超过 300KB 的 JPG。"""
    paths, failed = [], []
    for src in srcs:
        if os.path.splitext(src)[1].lower() != ".pdf":
            continue
        try:
            paths.extend(converter.pdf_to_small_jpgs(src, work_dir))
        except Exception as e:
            failed.append(os.path.basename(src) + f"（{e}）")
    return paths, failed


@app.post("/api/pdf-folder-to-jpg")
async def pdf_folder_to_jpg(files: list[UploadFile] = File(...)):
    """一键将上传文件夹中的所有 PDF 每页转换为小于 300KB 的 JPG。"""
    pdfs = [f for f in files if (f.filename or "").lower().endswith(".pdf")]
    if not pdfs:
        raise HTTPException(400, "所选文件夹中没有 PDF 文件")

    task_id = uuid.uuid4().hex[:12]
    task_upload = os.path.join(UPLOAD_DIR, task_id)
    task_output = os.path.join(OUTPUT_DIR, task_id)
    work_dir = os.path.join(task_output, "_jpgs")
    os.makedirs(task_upload, exist_ok=True)
    os.makedirs(work_dir, exist_ok=True)

    srcs, total_in, used_names = [], 0, set()
    for uf in pdfs:
        name = os.path.basename(uf.filename or "document.pdf")
        base, ext = os.path.splitext(name)
        suffix = 1
        while name in used_names:
            name = f"{base}_{suffix}{ext}"
            suffix += 1
        used_names.add(name)
        path = os.path.join(task_upload, name)
        file_size = 0
        with open(path, "wb") as f:
            while chunk := await uf.read(1024 * 1024):
                file_size += len(chunk)
                total_in += len(chunk)
                if file_size > MAX_FILE_SIZE:
                    f.close()
                    shutil.rmtree(task_upload, ignore_errors=True)
                    shutil.rmtree(task_output, ignore_errors=True)
                    raise HTTPException(413, f"文件 {name} 超过 {MAX_FILE_SIZE_MB}MB 限制")
                if total_in > MAX_BATCH_SIZE:
                    f.close()
                    shutil.rmtree(task_upload, ignore_errors=True)
                    shutil.rmtree(task_output, ignore_errors=True)
                    raise HTTPException(413, f"批量总大小超过 {MAX_BATCH_SIZE_MB}MB 限制")
                f.write(chunk)
        srcs.append(path)

    loop = asyncio.get_event_loop()
    try:
        images, failed = await loop.run_in_executor(executor, _pdf_folder_to_jpgs, srcs, work_dir)
    finally:
        shutil.rmtree(task_upload, ignore_errors=True)
    if not images:
        shutil.rmtree(task_output, ignore_errors=True)
        raise HTTPException(500, "PDF 转 JPG 全部失败: " + "; ".join(failed[:5]))

    out_path = os.path.join(task_output, f"pdf_to_jpg_under_300KB_{len(images)}images.zip")
    converter.zip_files(images, out_path)
    shutil.rmtree(work_dir, ignore_errors=True)
    _results[task_id] = out_path
    result = _build_result(task_id, out_path, total_in)
    result.update({"pdf_files": len(srcs), "image_files": len(images), "failed_files": failed})
    return result


@app.post("/api/scan-photo")
async def scan_photo(file: UploadFile = File(...)):
    """将一张拍摄的文档照片校正并增强为扫描 JPG。"""
    filename = os.path.basename(file.filename or "")
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    if ext not in converter.IMAGE_EXTS:
        raise HTTPException(400, "请上传 JPG、PNG、WEBP、HEIC 转换后的常见照片格式")

    task_id, src_path, task_output, src_size = await _save_upload(file)
    task_upload = os.path.dirname(src_path)
    base = os.path.splitext(os.path.basename(src_path))[0]
    out_path = os.path.join(task_output, f"{base}_scanned.jpg")
    loop = asyncio.get_event_loop()
    try:
        scan_info = await loop.run_in_executor(executor, converter.scan_photo, src_path, out_path)
    except Exception as e:
        shutil.rmtree(task_output, ignore_errors=True)
        raise HTTPException(500, f"扫描处理失败: {e}")
    finally:
        shutil.rmtree(task_upload, ignore_errors=True)

    _results[task_id] = out_path
    result = _build_result(task_id, out_path, src_size)
    result.update(scan_info)
    return result


# ---------------- 英文文献翻译 ----------------
def _trim_translate_tasks(limit: int = 50) -> None:
    """只保留最近的任务记录，避免长时间运行后内存无限增长。"""
    if len(_translate_tasks) <= limit:
        return
    for key in list(_translate_tasks)[:-limit]:
        _translate_tasks.pop(key, None)
        _translate_jobs.pop(key, None)


@app.post("/api/translate")
async def translate_document(
    file: UploadFile = File(...),
    output: str = Form("bilingual"),
    domain: str = Form(""),
    glossary: str = Form(""),
    translate_names: str = Form("0"),
    skip_references: str = Form("1"),
):
    """上传英文文献，后台翻译为中文 Word，通过进度接口轮询结果。

    翻译是多次网络调用，耗时可能达数分钟，因此立即返回 task_id；
    前端轮询 /api/translate/progress/{task_id} 获取进度与下载地址。
    """
    if output not in ("bilingual", "chinese", "both", "layout_pdf"):
        raise HTTPException(400, "output 必须是 bilingual / chinese / both / layout_pdf")
    if not TRANSLATE_CFG.configured:
        raise HTTPException(400, TRANSLATE_SETUP_HINT)

    cfg = TRANSLATE_CFG.with_overrides(
        domain=domain.strip() or None,
        glossary=_merged_glossary(glossary) or None,
        # 人名是否译为中文：默认保留原文拼音（按拼音猜汉字容易出错）
        translate_names=str(translate_names).lower() not in ("0", "false", "no", ""),
        skip_references=str(skip_references).lower() not in ("0", "false", "no"),
    )

    task_id, src_path, task_output, src_size = await _save_upload(file)
    task_upload = os.path.dirname(src_path)

    _translate_tasks[task_id] = {
        "state": "running",
        "done": 0,
        "total": 0,
        "message": "正在提取文献正文…",
        "src_size": src_size,
        "filename": os.path.basename(src_path),
    }
    _trim_translate_tasks()

    def progress(done: int, total: int, message: str) -> None:
        info = _translate_tasks.get(task_id)
        if info is not None:
            info["done"], info["total"], info["message"] = done, total, message

    def job() -> dict:
        # 源文件在翻译结束后才能删除，translator 全程需要读取它
        try:
            return translator.translate_document(
                src_path, task_output, cfg, output, progress=progress
            )
        finally:
            shutil.rmtree(task_upload, ignore_errors=True)

    async def runner() -> None:
        info = _translate_tasks[task_id]
        try:
            loop = asyncio.get_event_loop()
            stats = await loop.run_in_executor(translate_executor, job)
            out_path = stats["path"]
            _results[task_id] = out_path
            result = _build_result(task_id, out_path, src_size)
            result.update({
                "paragraphs": stats["paragraphs"],
                "translated": stats["translated"],
                "skipped_references": stats["skipped_references"],
            })
            # 原版式 PDF 会额外报告有多少段因版面重叠/放不下而保留原文
            for key in ("dropped", "unfitted"):
                if key in stats:
                    result[key] = stats[key]
            info.update(state="done", message="翻译完成", result=result)
        except Exception as exc:
            shutil.rmtree(task_output, ignore_errors=True)
            info.update(state="error", message=str(exc))
        finally:
            _translate_jobs.pop(task_id, None)

    _translate_jobs[task_id] = asyncio.create_task(runner())
    return {"task_id": task_id}


@app.get("/api/translate/progress/{task_id}")
async def translate_progress(task_id: str):
    """查询翻译进度；state=done 时 result 中带有下载地址。"""
    info = _translate_tasks.get(task_id)
    if not info:
        raise HTTPException(404, "翻译任务不存在或已过期")
    return {
        "state": info["state"],
        "done": info.get("done", 0),
        "total": info.get("total", 0),
        "message": info.get("message", ""),
        "result": info.get("result"),
    }


@app.get("/api/download/{task_id}")
async def download(task_id: str):
    path = _results.get(task_id)
    if not path or not os.path.isfile(path):
        raise HTTPException(404, "文件不存在或已过期")
    return FileResponse(path, filename=os.path.basename(path))


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


if __name__ == "__main__":
    import uvicorn
    # 0.0.0.0 = 监听所有网卡，局域网内其他电脑可通过本机 IP 访问
    uvicorn.run(app, host="0.0.0.0", port=8000)
