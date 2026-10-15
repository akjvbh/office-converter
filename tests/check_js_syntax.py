"""用 node 校验 static/app.js 的语法（比括号配平更权威）。

依赖 node，但不需要任何 npm 包。找不到 node 时以退出码 2 跳过。

用法：

    python tests/check_js_syntax.py
"""
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGETS = [os.path.join(BASE, "static", "app.js")]

node = shutil.which("node")
if node is None:
    print("找不到 node，跳过 JS 语法校验（装上 node 后可运行本脚本）")
    raise SystemExit(2)

print("node:", node)
version = subprocess.run([node, "--version"], capture_output=True, text=True)
print("version:", (version.stdout or version.stderr).strip())

failed = False
for target in TARGETS:
    # 浏览器按 ES module 加载 app.js，因此这里也按 module 解析
    with open(target, "r", encoding="utf-8") as handle:
        check = subprocess.run(
            [node, "--input-type=module", "--check"],
            stdin=handle, capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
    label = os.path.relpath(target, BASE)
    if check.returncode == 0:
        print(f"[OK] {label} 语法正确")
    else:
        failed = True
        print(f"[FAIL] {label}")
        print(check.stdout)
        print(check.stderr)

raise SystemExit(1 if failed else 0)
