"""
核心转换模块
支持的转换：
- Word/Excel/PPT -> PDF (通过 MS Office COM，Windows)
- PDF -> Word / 图片 / 文本
- 图片格式互转（PNG/JPG/BMP/WEBP/GIF/TIFF）、图片 OCR -> 可编辑 Word
- Excel <-> CSV
- Markdown -> HTML
- TXT -> PDF
"""
import os
import csv
import base64
import json
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
import pythoncom  # type: ignore


# ---------- PDF 合并与编辑 ----------
def merge_pdfs(paths: list, dst: str, mode: str = "sequence", insert_index: int = 0):
    """合并 PDF。sequence=顺序，alternate=逐页交替，prepend/append=插入到主 PDF 前/后。"""
    import fitz

    if len(paths) < 2:
        raise ValueError("至少需要两个 PDF 文件")
    docs = [fitz.open(path) for path in paths]
    output = fitz.open()
    try:
        if mode == "alternate":
            max_pages = max(len(doc) for doc in docs)
            for page_no in range(max_pages):
                for doc in docs:
                    if page_no < len(doc):
                        output.insert_pdf(doc, from_page=page_no, to_page=page_no)
        elif mode == "sequence":
            for doc in docs:
                output.insert_pdf(doc)
        elif mode in ("prepend", "append"):
            main, inserted = docs[0], docs[1:]
            if mode == "prepend":
                for doc in inserted:
                    output.insert_pdf(doc)
                output.insert_pdf(main)
            else:
                output.insert_pdf(main)
                for doc in inserted:
                    output.insert_pdf(doc)
        elif mode == "insert":
            main = docs[0]
            index = max(0, min(int(insert_index), len(main)))
            if index:
                output.insert_pdf(main, to_page=index - 1)
            for doc in docs[1:]:
                output.insert_pdf(doc)
            if index < len(main):
                output.insert_pdf(main, from_page=index)
        else:
            raise ValueError("不支持的合并模式")
        output.save(dst, garbage=4, deflate=True)
    finally:
        output.close()
        for doc in docs:
            doc.close()


def edit_pdf(src: str, dst: str, delete_pages: list[int] | None = None,
             rotations: dict | None = None, page_order: list[int] | None = None):
    """按 1-based 页码删除、旋转并重排 PDF 页面。"""
    import fitz

    source = fitz.open(src)
    result = fitz.open()
    delete = {int(page) - 1 for page in (delete_pages or [])}
    rotations = {int(k) - 1: int(v) for k, v in (rotations or {}).items()}
    order = [int(page) - 1 for page in page_order] if page_order else list(range(len(source)))
    if any(page < 0 or page >= len(source) for page in order):
        source.close()
        raise ValueError("页码超出范围")
    for page_no in order:
        if page_no in delete:
            continue
        result.insert_pdf(source, from_page=page_no, to_page=page_no)
        if page_no in rotations:
            result[-1].set_rotation((source[page_no].rotation + rotations[page_no]) % 360)
    if len(result) == 0:
        source.close()
        result.close()
        raise ValueError("不能删除全部页面")
    result.save(dst, garbage=4, deflate=True)
    result.close()
    source.close()


def annotate_pdf(src: str, dst: str, annotations: list | None = None):
    """将可视化编辑器提交的标注写入 PDF。坐标使用页面左上角原点的 PDF 坐标。"""
    import fitz
    doc = fitz.open(src)
    try:
        for item in annotations or []:
            page_no = int(item.get("page", 0))
            if page_no < 0 or page_no >= len(doc):
                continue
            page = doc[page_no]
            kind = item.get("type")
            if kind == "replace_text":
                rect = fitz.Rect(*[float(v) for v in item["rect"]])
                # 真正移除区域内原文字/图形对象，避免只是视觉覆盖后仍可复制旧内容。
                page.add_redact_annot(rect, fill=(1, 1, 1))
            elif kind == "text":
                rect = fitz.Rect(*[float(v) for v in item["rect"]])
                page.insert_textbox(rect, str(item.get("text", "")),
                                    fontsize=max(6, float(item.get("size", 14))),
                                    color=tuple(item.get("color", [0, 0, 0])),
                                    fontname="china-s", align=0)
            elif kind == "rect":
                rect = fitz.Rect(*[float(v) for v in item["rect"]])
                color = tuple(item.get("color", [0.1, 0.35, 0.85]))
                fill = item.get("fill")
                page.draw_rect(rect, color=color, fill=tuple(fill) if fill else None,
                               width=max(.5, float(item.get("width", 1.5))),
                               overlay=True, stroke_opacity=float(item.get("opacity", 1)),
                               fill_opacity=float(item.get("opacity", 1)))
            elif kind == "ellipse":
                rect = fitz.Rect(*[float(v) for v in item["rect"]])
                page.draw_oval(rect, color=tuple(item.get("color", [0.1, 0.35, 0.85])),
                               fill=tuple(item["fill"]) if item.get("fill") else None,
                               width=max(.5, float(item.get("width", 1.5))), overlay=True,
                               stroke_opacity=float(item.get("opacity", 1)),
                               fill_opacity=float(item.get("opacity", 1)))
            elif kind == "line":
                page.draw_line(fitz.Point(*item["start"]), fitz.Point(*item["end"]),
                               color=tuple(item.get("color", [0, 0, 0])),
                               width=max(0.5, float(item.get("width", 2))), overlay=True)
            elif kind == "arrow":
                start, end = item["start"], item["end"]
                page.draw_line(fitz.Point(*start), fitz.Point(*end),
                               color=tuple(item.get("color", [0.1, 0.35, 0.85])),
                               width=max(.5, float(item.get("width", 2))), overlay=True)
                page.draw_polyline([fitz.Point(*end), fitz.Point(end[0] - 8, end[1] - 4),
                                    fitz.Point(end[0] - 5, end[1] + 6), fitz.Point(*end)],
                                   color=tuple(item.get("color", [0.1, 0.35, 0.85])),
                                   width=max(.5, float(item.get("width", 2))), overlay=True)
            elif kind in ("highlight", "underline", "strike"):
                rect = fitz.Rect(*[float(v) for v in item["rect"]])
                color = tuple(item.get("color", [1, .75, .1]))
                if kind == "highlight":
                    page.draw_rect(rect, color=None, fill=color, overlay=True,
                                   fill_opacity=float(item.get("opacity", .35)))
                else:
                    y = rect.y1 - 2 if kind == "underline" else (rect.y0 + rect.y1) / 2
                    page.draw_line(fitz.Point(rect.x0, y), fitz.Point(rect.x1, y),
                                   color=color, width=max(1, float(item.get("width", 1.5))), overlay=True)
            elif kind == "image" and item.get("data"):
                raw = str(item["data"])
                if "," in raw:
                    raw = raw.split(",", 1)[1]
                page.insert_image(fitz.Rect(*[float(v) for v in item["rect"]]),
                                  stream=base64.b64decode(raw), keep_proportion=False,
                                  overlay=True, rotate=int(item.get("rotate", 0)))
            elif kind == "signature":
                for stroke in item.get("strokes", []):
                    for a, b in zip(stroke, stroke[1:]):
                        page.draw_line(fitz.Point(*a), fitz.Point(*b),
                                       color=(0.05, 0.15, 0.65), width=1.5, overlay=True)
        # 必须先应用所有删除标记，再写入替换文字；因此分两阶段处理。
        for page in doc:
            page.apply_redactions()
        for item in annotations or []:
            if item.get("type") != "replace_text":
                continue
            page_no = int(item.get("page", 0))
            if page_no < 0 or page_no >= len(doc):
                continue
            rect = fitz.Rect(*[float(v) for v in item["rect"]])
            doc[page_no].insert_textbox(
                rect, str(item.get("text", "")),
                fontsize=max(5, float(item.get("size", 11))),
                color=tuple(item.get("color", [0, 0, 0])),
                fontname="china-s", align=0,
            )
        doc.save(dst, garbage=4, deflate=True)
    finally:
        doc.close()

# ---------- 支持的转换映射 ----------
IMAGE_EXTS = {"png", "jpg", "jpeg", "jfif", "bmp", "webp", "gif", "tiff"}

CONVERSIONS = {
    "doc":  ["pdf", "docx", "txt"],
    "docx": ["pdf", "txt", "html"],
    "xls":  ["pdf", "xlsx", "csv"],
    "xlsx": ["pdf", "csv"],
    "csv":  ["xlsx"],
    "ppt":  ["pdf", "pptx"],
    "pptx": ["pdf", "png"],
    "pdf":  ["docx", "png", "jpg", "txt"],
    "md":   ["html", "pdf"],
    "txt":  ["pdf", "docx"],
    "png":  ["jpg", "jfif", "bmp", "webp", "gif", "tiff", "pdf", "ico", "docx"],
    "jpg":  ["png", "jfif", "bmp", "webp", "gif", "tiff", "pdf", "ico", "docx"],
    "jpeg": ["png", "jfif", "bmp", "webp", "gif", "tiff", "pdf", "ico", "docx"],
    "jfif": ["png", "jpg", "bmp", "webp", "gif", "tiff", "pdf", "ico", "docx"],
    "bmp":  ["png", "jpg", "jfif", "webp", "pdf", "docx"],
    "webp": ["png", "jpg", "jfif", "bmp", "pdf", "docx"],
    "gif":  ["png", "jpg", "jfif", "webp", "pdf", "docx"],
    "tiff": ["png", "jpg", "jfif", "pdf", "docx"],
}


def get_supported_targets(ext: str):
    return CONVERSIONS.get(ext.lower().lstrip("."), [])


# ---------- Office COM 转换（Windows） ----------
# 说明：Office 自动化会因为 Office 自动更新、激活/许可窗口、权限或没有可交互桌面而卡死。
# 这里统一加入「健康探测 + 超时 + 残留进程清理」，并在 Word 不可用且输入是 docx 时，
# 用内置排版引擎（docx_pdf.py）兜底出 PDF，避免一次异常就让请求永久挂住。
_OFFICE_IMAGE = {"word": "WINWORD.EXE", "excel": "EXCEL.EXE", "ppt": "POWERPNT.EXE"}
_OFFICE_LABEL = {"word": "Microsoft Word", "excel": "Microsoft Excel",
                 "ppt": "Microsoft PowerPoint"}
_OFFICE_STATE: dict = {}
_OFFICE_STATE_LOCK = threading.Lock()
_OFFICE_PROBE_LOCK = threading.Lock()
_ENGINE_NOTES: dict = {}


def _env_positive_int(name: str, default: int) -> int:
    raw = (os.environ.get(name) or "").strip()
    return int(raw) if raw.isdigit() and int(raw) > 0 else default


_OFFICE_TIMEOUT = _env_positive_int("OFFICE_TIMEOUT_SECONDS", 900)
_OFFICE_PROBE_TIMEOUT = _env_positive_int("OFFICE_PROBE_TIMEOUT_SECONDS", 45)
# 探测失败后隔多久允许再探测一次（秒）：用户处理后无需重启服务
_OFFICE_RECHECK_SECONDS = _env_positive_int("OFFICE_RECHECK_SECONDS", 600)


def _worker_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "office_worker.py")


def _no_window_flag() -> int:
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _worker_env() -> dict:
    """让子进程用 UTF-8 输出，避免中文 Windows 下按 GBK 解码报错、丢失错误详情。"""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _office_pids(app_name: str) -> set:
    """当前在跑的 Office 进程号（用于识别本次转换新起的实例）。"""
    image = _OFFICE_IMAGE.get(app_name)
    if not image or not sys.platform.startswith("win"):
        return set()
    try:
        completed = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=30, creationflags=_no_window_flag(),
        )
    except Exception:
        return set()
    pids = set()
    for line in (completed.stdout or "").splitlines():
        fields = [item.strip().strip('"') for item in line.split(",")]
        if len(fields) >= 2 and fields[0].lower().endswith(".exe"):
            try:
                pids.add(int(fields[1]))
            except ValueError:
                continue
    return pids


def _kill_process_tree(pid: int) -> None:
    try:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True,
                       timeout=30, creationflags=_no_window_flag())
    except Exception:
        pass


def _kill_orphan_office(app_name: str, before) -> int:
    """只清理本次转换期间新出现的 Office 进程，不碰用户自己打开的文档。"""
    orphans = _office_pids(app_name) - set(before or ())
    for pid in orphans:
        _kill_process_tree(pid)
    return len(orphans)


def _office_error_detail(output: str) -> str:
    for line in reversed((output or "").splitlines()):
        if line.strip():
            return line.strip()[:300]
    return ""


def _office_state_set(app_name: str, ok: bool, detail: str = "") -> None:
    with _OFFICE_STATE_LOCK:
        _OFFICE_STATE[app_name] = (ok, detail, time.time())


def _office_state_get(app_name: str):
    with _OFFICE_STATE_LOCK:
        return _OFFICE_STATE.get(app_name)


def reset_office_state() -> None:
    """清空 Office 可用性缓存（服务重启或用户处理后可用于重新探测）。"""
    with _OFFICE_STATE_LOCK:
        _OFFICE_STATE.clear()


def _record_engine_note(path: str, note: str) -> None:
    _ENGINE_NOTES[os.path.abspath(path)] = note
    print(f"[office] {note} -> {path}", flush=True)


def pop_engine_note(path: str) -> str:
    """取出并清除某个输出文件的转换方式说明（供前端提示）。"""
    return _ENGINE_NOTES.pop(os.path.abspath(path), "")


def _office_to_pdf(src: str, dst: str, app_name: str):
    """使用 MS Office COM 接口将 Word/Excel/PPT 转为 PDF"""
    import win32com.client
    pythoncom.CoInitialize()
    app = None
    document = None
    try:
        if app_name == "word":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Open(src, ReadOnly=True)
            document.SaveAs(dst, FileFormat=17)  # wdFormatPDF
        elif app_name == "excel":
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            document = app.Workbooks.Open(src, ReadOnly=True)
            document.ExportAsFixedFormat(0, dst)  # xlTypePDF
        elif app_name == "ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            document = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
            document.SaveAs(dst, 32)  # ppSaveAsPDF
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _office_convert_format(src: str, dst: str, app_name: str, fmt: int):
    """使用 Office COM 做格式另存（如 doc->docx, xls->xlsx, ppt->pptx, docx->txt）"""
    import win32com.client
    pythoncom.CoInitialize()
    app = None
    document = None
    try:
        if app_name == "word":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Open(src, ReadOnly=True)
            document.SaveAs(dst, FileFormat=fmt)
        elif app_name == "excel":
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            document = app.Workbooks.Open(src, ReadOnly=True)
            document.SaveAs(dst, FileFormat=fmt)
        elif app_name == "ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            document = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
            document.SaveAs(dst, fmt)
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _office_probe(dst: str, app_name: str) -> None:
    """空文档转 PDF：在隔离进程里判断 Office 自动化是否真的可用。"""
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    document = None
    try:
        if app_name == "word":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Add()
            document.SaveAs(dst, FileFormat=17)  # wdFormatPDF
        elif app_name == "excel":
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            document = app.Workbooks.Add()
            document.ExportAsFixedFormat(0, dst)  # xlTypePDF
        elif app_name == "ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            document = app.Presentations.Add()
            document.SaveAs(dst, 32)  # ppSaveAsPDF
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _probe_office(app_name: str) -> tuple:
    """在隔离子进程里跑一次空文档转换，返回 (是否可用, 原因)。"""
    label = _OFFICE_LABEL.get(app_name, "Office")
    probe_dir = tempfile.mkdtemp(prefix="office_probe_")
    dst = os.path.join(probe_dir, "probe.pdf")
    before = _office_pids(app_name)
    try:
        process = subprocess.Popen(
            [sys.executable, _worker_path(), "probe", "", dst, app_name],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            encoding="utf-8", errors="replace", env=_worker_env(),
            creationflags=_no_window_flag(),
        )
        try:
            stdout, stderr = process.communicate(timeout=_OFFICE_PROBE_TIMEOUT)
        except subprocess.TimeoutExpired:
            _kill_process_tree(process.pid)
            killed = _kill_orphan_office(app_name, before)
            return False, ("%s 自动化没有响应（%d 秒超时%s）"
                           % (label, _OFFICE_PROBE_TIMEOUT,
                              "，已结束残留进程" if killed else ""))
        if process.returncode != 0:
            _kill_orphan_office(app_name, before)
            detail = _office_error_detail(stderr or stdout)
            return False, ("%s 自动化不可用：%s" % (label, detail or "调用失败"))
        if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
            _kill_orphan_office(app_name, before)
            return False, ("%s 自动化未能输出文件" % label)
        return True, ""
    finally:
        shutil.rmtree(probe_dir, ignore_errors=True)


def _office_available(app_name: str) -> tuple:
    """返回 (是否可用, 原因)，结果在进程内缓存，避免每次转换都重新探测。

    探测失败后会隔一段时间再试（用户可能结束卡死的 Office 进程后恢复正常），
    这样不必重启服务就能自动回到 Office 通道。
    """
    cached = _office_state_get(app_name)
    if cached is not None:
        ok, detail, checked_at = cached
        if ok or (time.time() - checked_at) < _OFFICE_RECHECK_SECONDS:
            return ok, detail
        detail = "上次探测失败，重新探测：%s" % detail
    with _OFFICE_PROBE_LOCK:
        cached = _office_state_get(app_name)
        if cached is not None:
            ok, detail, checked_at = cached
            if ok or (time.time() - checked_at) < _OFFICE_RECHECK_SECONDS:
                return ok, detail
        label = _OFFICE_LABEL.get(app_name, "Office")
        if not sys.platform.startswith("win"):
            result = (False, f"{label} 自动化仅在 Windows 上可用")
        else:
            try:
                import win32com.client  # noqa: F401
                result = _probe_office(app_name)
            except Exception as exc:
                result = (False, f"无法调用 {label}：{exc}")
        _office_state_set(app_name, *result)
        return result


def _require_office(app_name: str) -> None:
    available, reason = _office_available(app_name)
    if not available:
        raise RuntimeError(reason)


def _run_office_isolated(operation: str, src: str, dst: str, app_name: str,
                         fmt: int | None = None, timeout: int | None = None):
    """在独立进程中运行 Office COM，避免 Word/Excel/PPT 异常拖垮 Web 服务。

    相比旧实现补上了超时与残留进程清理：Office 卡死时不再无限等待，也不会留下
    无响应的 WINWORD.EXE 影响后续转换。
    """
    command = [sys.executable, _worker_path(), operation, src, dst, app_name]
    if fmt is not None:
        command.append(str(fmt))

    limit = timeout or _OFFICE_TIMEOUT
    label = _OFFICE_LABEL.get(app_name, "Office")
    before = _office_pids(app_name)
    process = subprocess.Popen(
        command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        encoding="utf-8", errors="replace", env=_worker_env(),
        creationflags=_no_window_flag(),
    )
    try:
        stdout, stderr = process.communicate(timeout=limit)
    except subprocess.TimeoutExpired:
        _kill_process_tree(process.pid)
        killed = _kill_orphan_office(app_name, before)
        raise RuntimeError(
            "%s 在 %d 秒内没有响应，已中止本次转换%s。常见原因：Office 弹出了等待人工"
            "确认的窗口（激活/许可、只读提示等）或 Office 组件异常；可在任务管理器结束 "
            "%s 后重试。" % (label, limit, "，并清理了残留进程" if killed else "",
                             _OFFICE_IMAGE.get(app_name, label)))
    if process.returncode != 0:
        killed = _kill_orphan_office(app_name, before)
        detail = _office_error_detail(stderr or stdout)
        raise RuntimeError("%s 转换失败：%s%s" % (
            label, detail or "Office 子进程异常退出",
            "（已清理残留进程）" if killed else ""))
    if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
        _kill_orphan_office(app_name, before)
        raise RuntimeError("%s 转换未生成有效文件" % label)


def _office_to_pdf_with_fallback(src: str, dst: str, app_name: str) -> None:
    """Office → PDF：优先用 Office；Word 不可用且输入是 docx 时用内置排版引擎兜底。"""
    available, reason = _office_available(app_name)
    failure = None
    if available:
        try:
            _run_office_isolated("pdf", src, dst, app_name)
            return
        except RuntimeError as exc:
            failure = exc
            # 本次转换失败后直接标记不可用，避免后续请求重复等待超时
            _office_state_set(app_name, False, str(exc))
    else:
        failure = RuntimeError(reason)

    if app_name == "word" and src.lower().endswith(".docx"):
        try:
            import docx_pdf
            docx_pdf.convert(src, dst)
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError("%s；内置排版引擎也未能完成转换：%s" % (failure, exc))
        _record_engine_note(dst, "Word 不可用，已用内置排版引擎生成 PDF（版式与 Word 可能略有差异）")
        return
    if app_name == "word":
        raise RuntimeError(
            "%s；.doc 是旧版二进制格式，内置排版引擎无法解析，请先用 Word 另存为 .docx 再试。"
            % failure)
    raise RuntimeError(str(failure))


# ---------- PDF 转换 ----------
def pdf_to_docx(src: str, dst: str):
    from pdf2docx import Converter
    cv = Converter(src)
    cv.convert(dst)
    cv.close()


def pdf_to_images(src: str, out_dir: str, fmt: str = "png") -> list:
    """PDF 每页转图片，返回图片路径列表"""
    import fitz  # PyMuPDF
    doc = fitz.open(src)
    paths = []
    base = os.path.splitext(os.path.basename(src))[0]
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=150)
        p = os.path.join(out_dir, f"{base}_page{i + 1}.{fmt}")
        pix.save(p)
        paths.append(p)
    doc.close()
    return paths


def pdf_to_small_jpgs(src: str, out_dir: str, max_bytes: int = 300 * 1024) -> list:
    """将 PDF 每页转为 JPG，并将每张图片严格限制在指定大小以内。"""
    import fitz
    from io import BytesIO
    from PIL import Image

    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(src)
    base = os.path.splitext(os.path.basename(src))[0]
    paths = []
    for index, page in enumerate(doc, start=1):
        # 144 DPI 能保持大多数文字清晰；过大时由后续步骤逐级缩小。
        pix = page.get_pixmap(dpi=144, alpha=False)
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        best = b""
        candidate = image
        # 质量降至最低后仍超限时，持续缩小尺寸，确保结果不超过目标大小。
        while True:
            for quality in (82, 75, 68, 60, 52, 45, 38, 32, 25, 20, 15, 10, 5):
                data = BytesIO()
                candidate.save(data, "JPEG", quality=quality, optimize=True, progressive=True)
                best = data.getvalue()
                if len(best) <= max_bytes:
                    break
            if len(best) <= max_bytes:
                break
            if candidate.width == 1 and candidate.height == 1:
                raise ValueError("无法将页面压缩到目标大小")
            candidate = candidate.resize(
                (max(1, int(candidate.width * 0.8)), max(1, int(candidate.height * 0.8))),
                Image.Resampling.LANCZOS,
            )
        path = os.path.join(out_dir, f"{base}_page{index}.jpg")
        with open(path, "wb") as f:
            f.write(best)
        paths.append(path)
    doc.close()
    return paths


def pdf_to_txt(src: str, dst: str):
    import fitz
    doc = fitz.open(src)
    text = "\n".join(page.get_text() for page in doc)
    doc.close()
    with open(dst, "w", encoding="utf-8") as f:
        f.write(text)


# ---------- 图片转换 ----------
def _table_grid_from_image(image):
    """返回有完整边框表格的行、列边界；无法可靠识别时返回 None。"""
    import cv2
    import numpy as np

    gray = np.array(image.convert("L"))
    binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    height, width = binary.shape
    horizontal = cv2.morphologyEx(
        binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, width // 25), 1))
    )
    vertical = cv2.morphologyEx(
        binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, height // 25)))
    )

    def line_positions(mask, axis_length, is_vertical):
        coverage = (mask > 0).sum(axis=0 if is_vertical else 1)
        indexes = np.where(coverage >= axis_length * 0.45)[0]
        if len(indexes) == 0:
            return []
        groups = np.split(indexes, np.where(np.diff(indexes) > 1)[0] + 1)
        return [int(group.mean()) for group in groups if len(group)]

    columns = line_positions(vertical, height, True)
    rows = line_positions(horizontal, width, False)
    if len(columns) < 3 or len(rows) < 3:
        return None
    if any(b - a < 12 for a, b in zip(columns, columns[1:])):
        return None
    if any(b - a < 12 for a, b in zip(rows, rows[1:])):
        return None
    return rows, columns


def _save_scanned_editable_docx(src: str, dst: str, data: dict):
    """将扫描图作为底图，并以 Word 浮动文本框覆盖 OCR 到的文字区域。"""
    import win32com.client

    from PIL import Image

    original = Image.open(src)
    image_width, image_height = original.size
    page_width = 612  # Letter width in points; Word rescales the scan to fit this page.
    page_height = page_width * image_height / image_width
    scale = page_width / image_width
    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        document = word.Documents.Add()
        section = document.Sections(1)
        section.PageSetup.TopMargin = 0
        section.PageSetup.BottomMargin = 0
        section.PageSetup.LeftMargin = 0
        section.PageSetup.RightMargin = 0
        section.PageSetup.PageWidth = page_width
        section.PageSetup.PageHeight = page_height
        picture = document.InlineShapes.AddPicture(os.path.abspath(src), False, True)
        picture.Width = page_width
        picture.Height = page_height
        background = picture.ConvertToShape()
        background.RelativeHorizontalPosition = 1  # wdRelativeHorizontalPositionPage
        background.RelativeVerticalPosition = 1  # wdRelativeVerticalPositionPage
        background.Left = 0
        background.Top = 0
        background.WrapFormat.Type = 3  # wdWrapBehind

        for index, text in enumerate(data["text"]):
            text = text.strip()
            if not text or float(data["conf"][index]) < 0:
                continue
            left = data["left"][index] * scale
            top = data["top"][index] * scale
            width = max(3, data["width"][index] * scale)
            height = max(3, data["height"][index] * scale)
            box = document.Shapes.AddTextbox(1, left, top, width, height)
            box.RelativeHorizontalPosition = 1
            box.RelativeVerticalPosition = 1
            box.WrapFormat.Type = 3
            box.Fill.Visible = True
            box.Fill.ForeColor.RGB = 0xFFFFFF
            box.Line.Visible = False
            box.TextFrame.MarginLeft = 0
            box.TextFrame.MarginRight = 0
            box.TextFrame.MarginTop = 0
            box.TextFrame.MarginBottom = 0
            text_range = box.TextFrame.TextRange
            text_range.Text = text
            text_range.Font.Size = max(5, height * 0.85)
            text_range.ParagraphFormat.SpaceAfter = 0
        document.SaveAs2(os.path.abspath(dst), FileFormat=16)
        document.Close(False)
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def image_to_docx(src: str, dst: str):
    """将扫描图片导出为底图不变、文字可在 Word 中编辑的文档。"""
    import pytesseract
    from docx import Document
    from PIL import Image, ImageOps

    # 服务进程可能在安装 OCR 前启动，Windows 的 PATH 不会自动刷新。
    bundled_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.isfile(bundled_tesseract):
        pytesseract.pytesseract.tesseract_cmd = bundled_tesseract
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        raise RuntimeError("未检测到 Tesseract OCR，请按 README 安装后重试")

    image = Image.open(src).convert("RGB")
    # 灰度和自动对比度能提高扫描件、截图的文字识别率。
    image = ImageOps.autocontrast(ImageOps.grayscale(image))
    available_languages = set(pytesseract.get_languages(config=""))
    language = "chi_sim+eng" if {"chi_sim", "eng"} <= available_languages else "eng"
    data = pytesseract.image_to_data(
        image, lang=language, config="--psm 3", output_type=pytesseract.Output.DICT
    )
    # Word 生成文字图层同样可能卡死，放到隔离子进程里执行：超时会中止并给出提示，
    # 不会把 Web 服务的线程池拖住。
    _require_office("word")
    data_path = os.path.join(os.path.dirname(os.path.abspath(dst)), "_scan_text.json")
    with open(data_path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    try:
        _run_office_isolated("scan-docx", src, dst, "word", data_path)
    except Exception as exc:
        raise RuntimeError(f"无法生成扫描件可编辑图层：{exc}") from exc
    finally:
        try:
            os.remove(data_path)
        except OSError:
            pass


def image_convert(src: str, dst: str, target: str):
    from PIL import Image
    img = Image.open(src)
    target = target.lower()
    if target in ("jpg", "jpeg", "jfif"):
        img = img.convert("RGB")
        img.save(dst, "JPEG", quality=95)
    elif target == "pdf":
        img = img.convert("RGB")
        img.save(dst, "PDF")
    elif target == "ico":
        img.save(dst, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    else:
        img.save(dst)


# ---------- 照片扫描 ----------
def _order_document_corners(points):
    """将四个角排序为左上、右上、右下、左下。"""
    import numpy as np

    points = points.astype("float32")
    ordered = np.zeros((4, 2), dtype="float32")
    sums = points.sum(axis=1)
    diffs = np.diff(points, axis=1).reshape(-1)
    ordered[0] = points[np.argmin(sums)]
    ordered[2] = points[np.argmax(sums)]
    ordered[1] = points[np.argmin(diffs)]
    ordered[3] = points[np.argmax(diffs)]
    return ordered


def scan_photo(src: str, dst: str) -> dict:
    """将手机拍摄的纸质文件照片裁切、透视矫正并增强为扫描 JPG。"""
    import cv2
    import numpy as np

    image = cv2.imdecode(np.fromfile(src, dtype=np.uint8), cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError("无法读取图片文件")
    original_height, original_width = image.shape[:2]
    scale = min(1.0, 1600 / max(original_height, original_width))
    preview = cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(preview, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 50, 150)
    edges = cv2.dilate(edges, np.ones((3, 3), np.uint8), iterations=1)
    contours, _ = cv2.findContours(edges, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    min_area = preview.shape[0] * preview.shape[1] * 0.18
    corners = None
    for contour in sorted(contours, key=cv2.contourArea, reverse=True):
        if cv2.contourArea(contour) < min_area:
            break
        approximation = cv2.approxPolyDP(contour, 0.02 * cv2.arcLength(contour, True), True)
        if len(approximation) == 4 and cv2.isContourConvex(approximation):
            corners = approximation.reshape(4, 2).astype("float32") / scale
            break

    detected = corners is not None
    if detected:
        source = _order_document_corners(corners)
        width_top = np.linalg.norm(source[1] - source[0])
        width_bottom = np.linalg.norm(source[2] - source[3])
        height_right = np.linalg.norm(source[2] - source[1])
        height_left = np.linalg.norm(source[3] - source[0])
        width = max(1, int(max(width_top, width_bottom)))
        height = max(1, int(max(height_left, height_right)))
        target = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]], dtype="float32")
        transformed = cv2.warpPerspective(image, cv2.getPerspectiveTransform(source, target), (width, height))
    else:
        transformed = image

    gray = cv2.cvtColor(transformed, cv2.COLOR_BGR2GRAY)
    # 以较大模糊核估计光照背景，减轻手机拍照产生的阴影与色偏。
    background = cv2.GaussianBlur(gray, (0, 0), 35)
    normalized = cv2.divide(gray, background, scale=255)
    enhanced = cv2.normalize(normalized, None, 0, 255, cv2.NORM_MINMAX)
    enhanced = cv2.bilateralFilter(enhanced, 7, 40, 40)
    scanned = cv2.cvtColor(enhanced, cv2.COLOR_GRAY2BGR)
    # imwrite 在 Windows 上无法稳定处理中文路径，使用 tofile 写入。
    ok, encoded = cv2.imencode(".jpg", scanned, [cv2.IMWRITE_JPEG_QUALITY, 94])
    if not ok:
        raise ValueError("无法保存扫描结果")
    encoded.tofile(dst)
    return {"document_detected": detected, "width": scanned.shape[1], "height": scanned.shape[0]}


# ---------- Excel / CSV ----------
def xlsx_to_csv(src: str, dst: str):
    from openpyxl import load_workbook
    wb = load_workbook(src, read_only=True, data_only=True)
    ws = wb.active
    with open(dst, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            writer.writerow(["" if c is None else c for c in row])
    wb.close()


def csv_to_xlsx(src: str, dst: str):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    # 尝试多种编码
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            with open(src, "r", encoding=enc, newline="") as f:
                for row in csv.reader(f):
                    ws.append(row)
            break
        except UnicodeDecodeError:
            ws.delete_rows(1, ws.max_row)
            continue
    wb.save(dst)


# ---------- Markdown / TXT ----------
HTML_TEMPLATE = """<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>body{{max-width:800px;margin:40px auto;font-family:'Segoe UI','Microsoft YaHei',sans-serif;line-height:1.7;padding:0 20px;color:#333}}
pre{{background:#f5f5f5;padding:12px;border-radius:6px;overflow:auto}}
code{{background:#f5f5f5;padding:2px 5px;border-radius:3px}}
table{{border-collapse:collapse}}td,th{{border:1px solid #ccc;padding:6px 12px}}
blockquote{{border-left:4px solid #ddd;margin-left:0;padding-left:16px;color:#666}}</style>
</head><body>{content}</body></html>"""


def md_to_html(src: str, dst: str):
    import markdown
    with open(src, "r", encoding="utf-8") as f:
        html = markdown.markdown(f.read(), extensions=["tables", "fenced_code", "toc"])
    with open(dst, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE.format(content=html))


def text_to_pdf(src: str, dst: str):
    """TXT/MD 转 PDF（使用 PyMuPDF 排版，支持中文）"""
    import fitz
    with open(src, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()
    doc = fitz.open()
    page = doc.new_page()
    rect = fitz.Rect(50, 50, page.rect.width - 50, page.rect.height - 50)
    # 使用内置中文字体
    remaining = text
    while remaining:
        page = doc[-1]
        rc = page.insert_textbox(rect, remaining, fontsize=11,
                                 fontname="china-s", align=0)
        if rc >= 0:
            break
        # rc<0 表示文本溢出，粗略按比例截断换页
        # insert_textbox 不返回剩余文本，改用分行策略
        lines = remaining.split("\n")
        # 估算每页可容纳行数
        per_page = int(rect.height / (11 * 1.35))
        chunk = "\n".join(lines[:per_page])
        page.insert_textbox(rect, chunk, fontsize=11, fontname="china-s", align=0)
        remaining = "\n".join(lines[per_page:])
        if remaining:
            doc.new_page()
        else:
            break
    doc.save(dst)
    doc.close()


def docx_to_txt(src: str, dst: str):
    from docx import Document
    doc = Document(src)
    text = "\n".join(p.text for p in doc.paragraphs)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(text)


def docx_to_html(src: str, dst: str):
    from docx import Document
    doc = Document(src)
    parts = []
    for p in doc.paragraphs:
        style = p.style.name.lower() if p.style else ""
        if style.startswith("heading"):
            level = style.replace("heading", "").strip() or "2"
            level = level if level.isdigit() and 1 <= int(level) <= 6 else "2"
            parts.append(f"<h{level}>{p.text}</h{level}>")
        elif p.text.strip():
            parts.append(f"<p>{p.text}</p>")
    for table in doc.tables:
        rows = []
        for row in table.rows:
            cells = "".join(f"<td>{c.text}</td>" for c in row.cells)
            rows.append(f"<tr>{cells}</tr>")
        parts.append("<table>" + "".join(rows) + "</table>")
    with open(dst, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE.format(content="\n".join(parts)))


def txt_to_docx(src: str, dst: str):
    from docx import Document
    with open(src, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()
    doc = Document()
    for line in text.split("\n"):
        doc.add_paragraph(line)
    doc.save(dst)


def pptx_to_images(src: str, out_dir: str) -> list:
    """PPT 每页转 PNG（需要 PowerPoint）"""
    import win32com.client
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
        pres.SaveAs(out_dir, 18)  # ppSaveAsPNG，保存为文件夹
        pres.Close()
        app.Quit()
    finally:
        pythoncom.CoUninitialize()
    paths = []
    if os.path.isdir(out_dir):
        for name in sorted(os.listdir(out_dir)):
            paths.append(os.path.join(out_dir, name))
    return paths


def zip_files(paths: list, dst: str):
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in paths:
            zf.write(p, os.path.basename(p))


# ---------- 压缩功能 ----------
# 压缩级别: 80%(质量优先) / 60%(推荐) / 40%(体积优先) / 20%(强力)
COMPRESS_LEVELS = {"80", "60", "40", "20"}
# 可直接压缩的格式（其他格式走 zip 打包压缩）
COMPRESSIBLE_EXTS = IMAGE_EXTS | {"pdf"}

_IMG_QUALITY = {"80": 80, "60": 60, "40": 40, "20": 20}
_PDF_DPI = {"80": 150, "60": 110, "40": 80, "20": 60}
_PDF_JPG_Q = {"80": 80, "60": 60, "40": 40, "20": 20}


def compress_image(src: str, dst: str, level: str):
    """图片有损压缩（同时限制超大分辨率）"""
    from PIL import Image
    img = Image.open(src)
    ext = os.path.splitext(dst)[1].lower().lstrip(".")
    # 强力压缩时缩小超大图片
    max_side = {"80": 4096, "60": 2560, "40": 1600, "20": 1200}[level]
    if max(img.size) > max_side:
        img.thumbnail((max_side, max_side), Image.LANCZOS)
    q = _IMG_QUALITY[level]
    if ext in ("jpg", "jpeg", "jfif"):
        img.convert("RGB").save(dst, "JPEG", quality=q, optimize=True)
    elif ext == "png":
        # PNG 用调色板量化减小体积
        colors = {"80": 256, "60": 256, "40": 128, "20": 64}[level]
        img.convert("RGB").quantize(colors=colors).save(dst, "PNG", optimize=True)
    elif ext == "webp":
        img.save(dst, "WEBP", quality=q)
    else:
        img.save(dst, optimize=True)


def compress_pdf(src: str, dst: str, level: str):
    """
    PDF 压缩：
    - 80%: 保留较高质量
    - 60%/40%/20%: 逐步降低图片质量和分辨率
    """
    import fitz
    dpi = _PDF_DPI[level]
    q = _PDF_JPG_Q[level]
    src_doc = fitz.open(src)
    new_doc = fitz.open()
    for page in src_doc:
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("jpg", jpg_quality=q)
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=img_bytes)
    new_doc.save(dst, garbage=4, deflate=True)
    new_doc.close()
    src_doc.close()


def compress_file(src: str, out_dir: str, level: str = "60") -> str:
    """
    统一压缩入口，返回输出文件路径。
    图片/PDF 做有损/优化压缩，其他文件打包为 zip。
    """
    if level not in COMPRESS_LEVELS:
        raise ValueError(f"无效的压缩级别: {level}")
    ext = os.path.splitext(src)[1].lower().lstrip(".")
    base = os.path.splitext(os.path.basename(src))[0]

    if ext in IMAGE_EXTS:
        dst = os.path.join(out_dir, f"{base}_compressed.{ext}")
        compress_image(src, dst, level)
    elif ext == "pdf":
        dst = os.path.join(out_dir, f"{base}_compressed.pdf")
        compress_pdf(src, dst, level)
    else:
        dst = os.path.join(out_dir, f"{base}.zip")
        zip_files([src], dst)

    # 如果压缩后反而更大，直接复制原文件（zip 除外，格式已变）
    if os.path.getsize(dst) >= os.path.getsize(src) and not dst.endswith(".zip"):
        import shutil
        shutil.copy2(src, dst)
    return dst



# ---------- 统一入口 ----------
def convert(src: str, out_dir: str, target: str) -> str:
    """
    执行转换，返回输出文件路径。
    多文件输出（如 PDF 转图片）会自动打包为 zip。
    """
    ext = os.path.splitext(src)[1].lower().lstrip(".")
    base = os.path.splitext(os.path.basename(src))[0]
    target = target.lower()
    dst = os.path.join(out_dir, f"{base}.{target}")

    if target not in get_supported_targets(ext):
        raise ValueError(f"不支持从 {ext} 转换到 {target}")

    # 图片互转
    if ext in IMAGE_EXTS:
        if target == "docx":
            image_to_docx(src, dst)
        else:
            image_convert(src, dst, target)
        return dst

    # PDF 源
    if ext == "pdf":
        if target == "docx":
            pdf_to_docx(src, dst)
        elif target == "png":
            imgs = pdf_to_images(src, out_dir, target)
            if len(imgs) == 1:
                return imgs[0]
            dst = os.path.join(out_dir, f"{base}_images.zip")
            zip_files(imgs, dst)
        elif target == "jpg":
            imgs = pdf_to_small_jpgs(src, out_dir)
            if len(imgs) == 1:
                return imgs[0]
            dst = os.path.join(out_dir, f"{base}_images.zip")
            zip_files(imgs, dst)
        elif target == "txt":
            pdf_to_txt(src, dst)
        return dst

    # Word
    if ext in ("doc", "docx"):
        if target == "pdf":
            # Word 不可用时，docx 走内置排版引擎兜底，保证能拿到 PDF
            _office_to_pdf_with_fallback(src, dst, "word")
        elif target == "docx":
            _require_office("word")
            _run_office_isolated("format", src, dst, "word", 16)  # wdFormatDocumentDefault
        elif target == "txt":
            if ext == "docx":
                docx_to_txt(src, dst)
            else:
                _require_office("word")
                _run_office_isolated("format", src, dst, "word", 7)  # wdFormatEncodedText
        elif target == "html":
            docx_to_html(src, dst)
        return dst

    # Excel
    if ext in ("xls", "xlsx"):
        if target == "pdf":
            _office_to_pdf_with_fallback(src, dst, "excel")
        elif target == "xlsx":
            _require_office("excel")
            _run_office_isolated("format", src, dst, "excel", 51)  # xlOpenXMLWorkbook
        elif target == "csv":
            if ext == "xlsx":
                xlsx_to_csv(src, dst)
            else:
                _require_office("excel")
                _run_office_isolated("format", src, dst, "excel", 6)  # xlCSV
        return dst

    if ext == "csv" and target == "xlsx":
        csv_to_xlsx(src, dst)
        return dst

    # PPT
    if ext in ("ppt", "pptx"):
        if target == "pdf":
            _office_to_pdf_with_fallback(src, dst, "ppt")
        elif target == "pptx":
            _require_office("ppt")
            _run_office_isolated("format", src, dst, "ppt", 24)  # ppSaveAsOpenXMLPresentation
        elif target == "png":
            img_dir = os.path.join(out_dir, f"{base}_slides")
            imgs = pptx_to_images(src, img_dir)
            dst = os.path.join(out_dir, f"{base}_slides.zip")
            zip_files(imgs, dst)
        return dst

    # Markdown
    if ext == "md":
        if target == "html":
            md_to_html(src, dst)
        elif target == "pdf":
            text_to_pdf(src, dst)
        return dst

    # TXT
    if ext == "txt":
        if target == "pdf":
            text_to_pdf(src, dst)
        elif target == "docx":
            txt_to_docx(src, dst)
        return dst

    raise ValueError(f"未实现的转换：{ext} -> {target}")
