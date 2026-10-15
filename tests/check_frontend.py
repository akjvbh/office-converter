"""前端静态检查：不需要 node 或浏览器，只依赖标准库。

检查内容：

1. static/app.js 的括号与引号是否配平（跳过注释、字符串、模板串、正则）
2. static/index.html 里引用的静态资源是否存在
3. app.js 里 $('id') 引用的 id 是否都在 index.html 中出现
4. index.html 里的 id 是否有重复
5. app.css 的花括号是否配平
6. index.html 标签是否闭合

用法：

    python tests/check_frontend.py

退出码非 0 表示发现问题。
"""
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATIC = os.path.join(BASE, "static")

problems = []


def read(path):
    with open(path, "r", encoding="utf-8") as handle:
        return handle.read()


# ---------- 1. JS 定界符配平 ----------
def check_js_delimiters(src, label):
    """逐字符扫描，跳过注释/字符串/模板串/正则，检查 ()[]{} 是否配平。"""
    pairs = {")": "(", "]": "[", "}": "{"}
    stack = []
    i, n = 0, len(src)
    line = 1
    prev_significant = ""   # 上一个有意义的字符，用来判断 / 是除号还是正则
    while i < n:
        ch = src[i]
        if ch == "\n":
            line += 1
            i += 1
            continue

        # 注释
        if ch == "/" and i + 1 < n and src[i + 1] == "/":
            i = src.find("\n", i)
            if i < 0:
                break
            continue
        if ch == "/" and i + 1 < n and src[i + 1] == "*":
            end = src.find("*/", i + 2)
            if end < 0:
                problems.append(f"{label}: 第 {line} 行有未闭合的块注释")
                return
            line += src.count("\n", i, end)
            i = end + 2
            continue

        # 字符串 / 模板串
        if ch in "\"'`":
            quote = ch
            j = i + 1
            while j < n:
                if src[j] == "\\":
                    j += 2
                    continue
                if src[j] == "\n":
                    if quote != "`":
                        problems.append(f"{label}: 第 {line} 行字符串未闭合")
                        return
                    line += 1
                if src[j] == quote:
                    break
                # 模板串里的 ${...} 需要再平衡检查，这里简单跳过整段
                j += 1
            if j >= n:
                problems.append(f"{label}: 第 {line} 行字符串/模板串未闭合")
                return
            i = j + 1
            prev_significant = quote
            continue

        # 正则字面量（粗略判断：/ 出现在表达式位置）
        if ch == "/" and prev_significant in ("", "(", ",", "=", ":", "[", "!", "&", "|", "?", "{", "}", ";"):
            j = i + 1
            in_class = False
            while j < n:
                if src[j] == "\\":
                    j += 2
                    continue
                if src[j] == "[":
                    in_class = True
                elif src[j] == "]":
                    in_class = False
                elif src[j] == "/" and not in_class:
                    break
                elif src[j] == "\n":
                    break
                j += 1
            i = j + 1
            continue

        if ch in "([{":
            stack.append((ch, line))
        elif ch in ")]}":
            if not stack:
                problems.append(f"{label}: 第 {line} 行多出一个 '{ch}'")
                return
            opener, opened_at = stack.pop()
            if opener != pairs[ch]:
                problems.append(f"{label}: 第 {line} 行的 '{ch}' 与第 {opened_at} 行的 '{opener}' 不匹配")
                return
        if not ch.isspace():
            prev_significant = ch
        i += 1

    if stack:
        opener, opened_at = stack[-1]
        problems.append(f"{label}: 第 {opened_at} 行的 '{opener}' 没有闭合（还有 {len(stack)} 个未闭合）")
    else:
        print(f"  [OK] {label} 定界符全部配平")


# ---------- 2. CSS 花括号配平 ----------
def check_css_braces(src, label):
    stripped = re.sub(r"/\*.*?\*/", "", src, flags=re.S)
    stripped = re.sub(r"\"[^\"\n]*\"|'[^'\n]*'", "", stripped)
    depth = 0
    for index, ch in enumerate(stripped):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth < 0:
                line = stripped.count("\n", 0, index) + 1
                problems.append(f"{label}: 第 {line} 行多出一个 '}}'")
                return
    if depth:
        problems.append(f"{label}: 还有 {depth} 个 '{{' 没有闭合")
    else:
        print(f"  [OK] {label} 花括号配平")


# ---------- 3. HTML 检查 ----------
VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link",
             "meta", "param", "source", "track", "wbr"}


def check_html(html, label):
    # id 重复
    ids = re.findall(r'\bid="([^"]+)"', html)
    duplicates = {value for value in ids if ids.count(value) > 1}
    if duplicates:
        problems.append(f"{label}: id 重复 -> {sorted(duplicates)}")
    else:
        print(f"  [OK] {label} 没有重复的 id（共 {len(ids)} 个）")

    # 标签配平（跳过 script/style 内容）
    body = re.sub(r"<(script|style)\b.*?</\1>", "", html, flags=re.S | re.I)
    body = re.sub(r"<!--.*?-->", "", body, flags=re.S)
    body = re.sub(r"<!DOCTYPE[^>]*>", "", body, flags=re.I)
    stack = []
    for match in re.finditer(r"<(/?)([a-zA-Z][\w-]*)([^>]*?)(/?)>", body):
        closing, tag, attrs, self_close = match.group(1), match.group(2).lower(), match.group(3), match.group(4)
        line = body.count("\n", 0, match.start()) + 1
        if tag in VOID_TAGS or self_close:
            continue
        if closing:
            if not stack:
                problems.append(f"{label}: 第 {line} 行多出 </{tag}>")
                return
            opener, opened_at = stack.pop()
            if opener != tag:
                problems.append(f"{label}: 第 {line} 行的 </{tag}> 与第 {opened_at} 行的 <{opener}> 不匹配")
                return
        else:
            stack.append((tag, line))
    if stack:
        opener, opened_at = stack[-1]
        problems.append(f"{label}: 第 {opened_at} 行的 <{opener}> 没有闭合")
    else:
        print(f"  [OK] {label} 标签全部闭合")
    return ids


# ---------- 主流程 ----------
print("检查前端资源：")

js_path = os.path.join(STATIC, "app.js")
css_path = os.path.join(STATIC, "app.css")
html_path = os.path.join(STATIC, "index.html")
for path in (js_path, css_path, html_path):
    if not os.path.isfile(path):
        problems.append(f"缺少文件: {path}")
        print(f"  [!!] 缺少文件: {path}")

if not problems:
    js = read(js_path)
    css = read(css_path)
    html = read(html_path)

    check_js_delimiters(js, "app.js")
    check_css_braces(css, "app.css")
    ids = check_html(html, "index.html")

    # 资源引用
    for ref in re.findall(r'(?:src|href)="(/static/[^"?]+)', html):
        target = os.path.join(STATIC, os.path.basename(ref))
        if os.path.isfile(target):
            print(f"  [OK] 静态资源存在: {ref}")
        else:
            problems.append(f"index.html 引用了不存在的资源: {ref}")

    # JS 里用到的 id 是否都存在于 HTML
    used = set(re.findall(r"\$\('([A-Za-z][\w-]*)'\)", js))
    missing = sorted(used - set(ids))
    if missing:
        problems.append(f"app.js 引用了 index.html 中不存在的 id: {missing}")
    else:
        print(f"  [OK] app.js 引用的 {len(used)} 个 id 在 index.html 中都存在")

    # 选择器引用的容器是否存在
    for selector in re.findall(r"\$\$\('#([\w-]+)", js):
        if selector not in ids:
            problems.append(f"app.js 用了 #{selector} 作为选择器前缀，但该 id 不存在")
    for selector in re.findall(r"\$\('#([\w-]+) ", js):
        if selector not in ids:
            problems.append(f"app.js 用了 #{selector} 作为选择器前缀，但该 id 不存在")

print()
if problems:
    print("发现问题：")
    for item in problems:
        print("  -", item)
    sys.exit(1)
print("全部检查通过。")
