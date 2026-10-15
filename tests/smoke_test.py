"""真实端到端冒烟测试：启动服务后走一遍关键接口。

需要项目虚拟环境内的 fastapi / uvicorn / PyMuPDF。

用法：

    .venv-ocr\\Scripts\\python.exe tests/smoke_test.py

会临时占用 8123 端口，结束后自动关闭。
"""
import io
import json
import os
import shutil
import stat
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = 8123
ROOT = f"http://127.0.0.1:{PORT}"

failures = []


def make_workdir(prefix):
    """建一个临时工作目录；系统临时目录不可写时改在仓库内建，并保证清理。"""
    chosen = tempfile.gettempdir()
    try:
        if os.path.commonpath([os.path.abspath(chosen), BASE]) == BASE:
            chosen = os.path.join(BASE, "_tmp_tests")
            os.makedirs(chosen, exist_ok=True)
    except ValueError:
        pass
    return tempfile.mkdtemp(prefix=prefix, dir=chosen)


def remove_workdir(path):
    """尽力删除临时目录。

    Windows 上刚结束的子进程可能还短暂持有句柄，所以先重试几次，
    并顺手清掉只读属性；实在删不掉也不影响测试结果。
    """

    def on_error(func, target, _exc):
        try:
            os.chmod(target, stat.S_IWRITE)
            func(target)
        except OSError:
            pass

    for _ in range(5):
        if not os.path.exists(path):
            return
        shutil.rmtree(path, onerror=on_error)
        if not os.path.exists(path):
            return
        time.sleep(0.4)


def check(label, condition, detail=""):
    if condition:
        print(f"  [OK] {label}")
    else:
        print(f"  [FAIL] {label} {detail}")
        failures.append(label)


def get(path, raw=False):
    with urllib.request.urlopen(ROOT + path, timeout=30) as resp:
        body = resp.read()
        return resp.status, body if raw else body.decode("utf-8", "replace")


def post_multipart(path, fields, files, timeout=120):
    """极简 multipart 编码，避免依赖 requests。"""
    boundary = "----smoketest" + str(int(time.time() * 1000))
    chunks = []
    for name, value in fields.items():
        chunks.append(
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"\r\n\r\n{value}\r\n".encode()
        )
    for name, (filename, content) in files:
        chunks.append(
            f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"; "
            f"filename=\"{filename}\"\r\nContent-Type: application/octet-stream\r\n\r\n".encode()
        )
        chunks.append(content)
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode())
    body = b"".join(chunks)
    req = urllib.request.Request(
        ROOT + path, data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as exc:
        return exc.code, json.loads(exc.read().decode("utf-8", "replace"))


def make_pdf(path, labels):
    import fitz
    doc = fitz.open()
    for label in labels:
        page = doc.new_page(width=300, height=400)
        page.insert_text((60, 200), label, fontsize=24)
    doc.save(path)
    doc.close()


def main():
    work = make_workdir("smoke-")
    try:
        return checks(work)
    finally:
        remove_workdir(work)


def checks(work):
    src = os.path.join(work, "sample.pdf")
    make_pdf(src, ["ALPHA", "BETA", "GAMMA"])

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    # 不用文件当子进程的 stdout：Windows 上子进程会一直持有该文件句柄，
    # 即使进程已结束，删除工作目录时也会报 WinError 32。
    # 改成在后台线程里持续读取管道，输出留在内存中，只有启动失败时才打印。
    proc = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "main:app", "--host", "127.0.0.1",
         "--port", str(PORT), "--log-level", "warning"],
        cwd=BASE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        env=env, text=True, encoding="utf-8", errors="replace",
    )
    output = []
    reader = threading.Thread(target=lambda: [output.append(line) for line in proc.stdout], daemon=True)
    reader.start()

    def server_log():
        reader.join(timeout=3)
        return "".join(output)[-3000:] or "(服务没有输出)"

    try:
        # 等服务就绪
        ready = False
        for _ in range(80):
            if proc.poll() is not None:
                break
            try:
                get("/api/formats")
                ready = True
                break
            except Exception:
                time.sleep(0.25)
        if not ready:
            print(server_log())
            print("服务未能启动")
            return 1

        print("服务已就绪，开始检查接口：")

        # 1. 页面与静态资源
        status, html = get("/")
        check("GET / 返回 200", status == 200)
        check("index.html 引用了 app.css", "/static/app.css" in html)
        check("index.html 引用了 app.js", "/static/app.js" in html)
        check("标题已更新", "办公文档转换器" in html)
        status, css = get("/static/app.css", raw=True)
        check("GET /static/app.css 返回 200", status == 200)
        status, js = get("/static/app.js", raw=True)
        check("GET /static/app.js 返回 200", status == 200)

        # 2. 配置接口
        status, body = get("/api/config")
        config = json.loads(body)
        check("GET /api/config 含上传上限", "max_file_size_mb" in config and "max_batch_size_mb" in config)
        check("GET /api/config 含翻译开关", "translate_configured" in config)

        # 3. 格式映射
        status, body = get("/api/formats")
        formats = json.loads(body)
        check("GET /api/formats 是映射表", isinstance(formats, dict) and "docx" in formats)

        with open(src, "rb") as handle:
            pdf_bytes = handle.read()

        # 4. 压缩（顺带验证前端压缩预估依赖的接口）
        for level in ("80", "20"):
            status, data = post_multipart("/api/compress", {"level": level}, [("file", ("sample.pdf", pdf_bytes))])
            check(f"POST /api/compress level={level}", status == 200 and data.get("output_size", 0) > 0,
                  f"status={status} data={data}")

        # 5. PDF 编辑：删除第 2 页
        status, data = post_multipart(
            "/api/pdf-edit",
            {"delete_pages": "2", "rotations": "{}", "page_order": "", "annotations": "[]"},
            [("file", ("sample.pdf", pdf_bytes))],
        )
        check("POST /api/pdf-edit 删除页", status == 200 and data.get("output_size", 0) > 0,
              f"status={status} data={data}")
        if status == 200:
            out = urllib.request.urlopen(ROOT + data["download_url"], timeout=30).read()
            import fitz
            doc = fitz.open(stream=out, filetype="pdf")
            texts = [page.get_text().strip() for page in doc]
            doc.close()
            check("删除第 2 页后的内容", texts == ["ALPHA", "GAMMA"], str(texts))

        # 6. PDF 编辑：重排 + 标注（验证前端提交的 page 语义）
        status, data = post_multipart(
            "/api/pdf-edit",
            {"delete_pages": "", "rotations": "{}", "page_order": "3,1,2",
             "annotations": json.dumps([{"type": "rect", "page": 0,
                                          "rect": [20, 20, 120, 80], "color": [1, 0, 0]}])},
            [("file", ("sample.pdf", pdf_bytes))],
        )
        check("POST /api/pdf-edit 重排+标注", status == 200, f"status={status} data={data}")
        if status == 200:
            out = urllib.request.urlopen(ROOT + data["download_url"], timeout=30).read()
            import fitz
            doc = fitz.open(stream=out, filetype="pdf")
            texts = [page.get_text().strip() for page in doc]
            doc.close()
            check("重排 page_order=3,1,2 后的顺序", texts == ["GAMMA", "ALPHA", "BETA"], str(texts))

        # 7. PDF 预览
        status, data = post_multipart("/api/pdf-preview", {}, [("file", ("sample.pdf", pdf_bytes))])
        check("POST /api/pdf-preview", status == 200 and len(data.get("pages", [])) == 3,
              f"status={status}")

        # 8. PDF 合并
        status, data = post_multipart(
            "/api/pdf-merge",
            {"mode": "sequence", "insert_index": "0"},
            [("files", ("a.pdf", pdf_bytes)), ("files", ("b.pdf", pdf_bytes))],
        )
        check("POST /api/pdf-merge", status == 200 and data.get("output_size", 0) > 0,
              f"status={status} data={data}")

        # 9. 批量转换
        status, data = post_multipart(
            "/api/batch",
            {"mode": "compress", "level": "60"},
            [("files", ("a.pdf", pdf_bytes)), ("files", ("b.pdf", pdf_bytes))],
        )
        check("POST /api/batch 压缩两个文件", status == 200 and data.get("success_files") == 2,
              f"status={status} data={data}")

        # 10. 错误处理：不支持的目标格式
        status, data = post_multipart("/api/convert", {"target": "exe"},
                                      [("file", ("sample.pdf", pdf_bytes))])
        check("不支持的目标格式返回 400", status == 400, f"status={status} data={data}")

        # 11. 错误处理：不存在的下载任务
        try:
            get("/api/download/doesnotexist")
            check("不存在的下载任务返回 404", False)
        except urllib.error.HTTPError as exc:
            check("不存在的下载任务返回 404", exc.code == 404, f"status={exc.code}")

    finally:
        proc.terminate()
        try:
            proc.wait(timeout=15)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=10)
        if proc.stdout:
            proc.stdout.close()
        reader.join(timeout=3)

    print()
    if failures:
        print(f"有 {len(failures)} 项失败：{failures}")
        print("---- 服务输出 ----")
        print(server_log())
        return 1
    print("全部冒烟检查通过。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
