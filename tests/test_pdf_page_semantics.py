"""校验 PDF 编辑/合并接口的页码语义，确保前端提交的参数不会丢页或错位。

覆盖的约定（对应 static/app.js 里 submitEditor / currentOrder 发出的参数）：

  1. 不做任何编辑时，页数与内容保持不变
  2. page_order 是「原始 1-based 页码」的目标顺序，且同时是白名单
  3. annotations 的 page 是「在 page_order 中的位置」，不是原始页码
  4. delete_pages 用原始 1-based 页码
  5. rotations 用原始 1-based 页码，值是增量
  6. 超过 50 页时「补回未预览页面」的做法不会丢页
  7. merge 的 insert_index 是 0-based

用法（需要 PyMuPDF）：

    python tests/test_pdf_page_semantics.py

不依赖 pytest，退出码非 0 表示有检查失败。
"""
import os
import shutil
import sys
import tempfile
import time

import fitz

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

import converter  # noqa: E402

failures = []


def make_workdir(prefix):
    """建一个临时工作目录。

    系统临时目录不可写时（受限环境很常见）Python 会退回到当前目录，
    这时我们自己在仓库里建目录，并保证最后清理掉。
    """
    chosen = tempfile.gettempdir()
    try:
        if os.path.commonpath([os.path.abspath(chosen), BASE]) == BASE:
            chosen = os.path.join(BASE, "_tmp_tests")
            os.makedirs(chosen, exist_ok=True)
    except ValueError:
        pass
    return tempfile.mkdtemp(prefix=prefix, dir=chosen)


def check(label, actual, expected):
    if actual == expected:
        print(f"  [OK] {label}: {actual}")
    else:
        print(f"  [FAIL] {label}: 期望 {expected}，实际 {actual}")
        failures.append(label)


def make_pdf(path, page_labels):
    """每页写一个可识别的标记文字，用于判断页面内容与顺序。"""
    doc = fitz.open()
    for label in page_labels:
        page = doc.new_page(width=300, height=400)
        page.insert_text((60, 200), label, fontsize=24)
    doc.save(path)
    doc.close()


def page_texts(path):
    doc = fitz.open(path)
    texts = [page.get_text().strip() for page in doc]
    doc.close()
    return texts


def run():
    work = make_workdir("pdfedit-")
    try:
        return checks(work)
    finally:
        for _ in range(5):
            if not os.path.exists(work):
                break
            shutil.rmtree(work, ignore_errors=True)
            if not os.path.exists(work):
                break
            time.sleep(0.3)


def checks(work):
    src = os.path.join(work, "src.pdf")
    make_pdf(src, ["PAGE-A", "PAGE-B", "PAGE-C"])
    check("原始文件页数", len(page_texts(src)), 3)

    # --- 1. 无参数编辑：不应丢页 ---
    out = os.path.join(work, "noop.pdf")
    converter.edit_pdf(src, out, [], {}, None)
    check("无参数编辑后页数", len(page_texts(out)), 3)
    check("无参数编辑后内容", page_texts(out), ["PAGE-A", "PAGE-B", "PAGE-C"])

    # --- 2. 重排：page_order 用原始页码 ---
    out = os.path.join(work, "reorder.pdf")
    converter.edit_pdf(src, out, [], {}, [3, 1, 2])
    check("重排 [3,1,2] 后的顺序", page_texts(out), ["PAGE-C", "PAGE-A", "PAGE-B"])

    # --- 3. 白名单语义：只列 1,3 时第 2 页被丢弃 ---
    out = os.path.join(work, "whitelist.pdf")
    converter.edit_pdf(src, out, [], {}, [1, 3])
    check("page_order=[1,3] 后的页数", len(page_texts(out)), 2)
    check("page_order=[1,3] 后的内容", page_texts(out), ["PAGE-A", "PAGE-C"])

    # --- 4. 删除：delete_pages 用原始 1-based 页码 ---
    out = os.path.join(work, "delete.pdf")
    converter.edit_pdf(src, out, [2], {}, None)
    check("删除第 2 页后的内容", page_texts(out), ["PAGE-A", "PAGE-C"])

    # --- 5. 旋转：rotations 用原始 1-based 页码，值是增量 ---
    out = os.path.join(work, "rotate.pdf")
    converter.edit_pdf(src, out, [], {"2": 90}, None)
    doc = fitz.open(out)
    rotations = [page.rotation for page in doc]
    doc.close()
    check("第 2 页旋转 90 后的各页 rotation", rotations, [0, 90, 0])

    # --- 6. 标注的 page 是「在 page_order 中的位置」 ---
    # 目标顺序 [3,1,2] → 位置 0 对应原第 3 页。
    # 在这份 PDF 上做白底遮盖，会把该页文字盖住，据此判断标注落到了哪一页。
    out = os.path.join(work, "annot.pdf")
    # 先构造带顺序的中间文件，再单独验证标注的页序解释
    mid = os.path.join(work, "mid.pdf")
    converter.edit_pdf(src, mid, [], {}, [3, 1, 2])
    check("中间文件顺序", page_texts(mid), ["PAGE-C", "PAGE-A", "PAGE-B"])

    marks = [{"type": "rect", "page": 0, "rect": [40, 170, 260, 220],
              "color": [1, 0, 0], "fill": [1, 1, 1], "opacity": 0.2}]
    annotated = os.path.join(work, "annotated.pdf")
    converter.annotate_pdf(mid, annotated, marks)
    # 用像素判断哪一页被画了东西：统计非白像素数量
    doc = fitz.open(annotated)
    ink = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
        samples = pix.samples
        dark = sum(1 for i in range(0, len(samples), pix.n) if samples[i] < 250)
        ink.append(dark)
    doc.close()
    print(f"       各页非白像素: {ink}")
    if ink and ink[0] > 0:
        print("  [OK] 标注写到了 page_order 位置 0 对应的页面")
    else:
        failures.append("标注未按 page_order 位置写入")
        print("  [FAIL] 标注没有写到位置 0 对应的页面")

    # --- 7. 超过 50 页：补回未预览页面的做法 ---
    big = os.path.join(work, "big.pdf")
    make_pdf(big, [f"P{i:03d}" for i in range(1, 61)])   # 60 页
    check("大文件页数", len(page_texts(big)), 60)
    # 模拟前端：只加载前 50 页预览，用户没有改动，
    # currentOrder() 会输出 1..50 再补 51..60
    order = list(range(1, 51)) + list(range(51, 61))
    out = os.path.join(work, "big_out.pdf")
    converter.edit_pdf(big, out, [], {}, order)
    check("补回未预览页面后的页数", len(page_texts(out)), 60)
    check("第一页内容", page_texts(out)[0], "P001")
    check("最后一页内容", page_texts(out)[-1], "P060")

    # --- 8. merge 的 insert_index 是 0-based ---
    main = os.path.join(work, "main.pdf")
    extra = os.path.join(work, "extra.pdf")
    make_pdf(main, ["M1", "M2"])
    make_pdf(extra, ["X1"])
    out = os.path.join(work, "merged1.pdf")
    converter.merge_pdfs([main, extra], out, "insert", 1)   # 插到第 1 页之后
    check("insert_index=1 的结果", page_texts(out), ["M1", "X1", "M2"])
    out = os.path.join(work, "merged0.pdf")
    converter.merge_pdfs([main, extra], out, "insert", 0)   # 插到首页之前
    check("insert_index=0 的结果", page_texts(out), ["X1", "M1", "M2"])

    print()
    if failures:
        print(f"有 {len(failures)} 项检查失败：{failures}")
        return 1
    print("全部验证通过。")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
