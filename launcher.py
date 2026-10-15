"""服务启动入口（由 start.bat 调用）。

为什么不让批处理直接启动服务：

批处理文件里只要同时出现「非 ASCII 文本（中文）」和「LF 换行」，cmd 就会错位解析，
把命令拆成乱码执行，窗口一闪就退出、服务根本没起来。这里把中文提示、日志与错误处理
都放到 Python 里，start.bat 只保留最简的纯 ASCII 逻辑，从根上避免这个问题。

职责：
- 打印启动提示，并把服务输出同时显示在窗口、写入 _start_log.txt
- python 缺失、main.py 缺失、启动异常时给出明确中文提示，不静默退出
"""
import os
import subprocess
import sys
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(BASE_DIR, "_start_log.txt")
PYTHON = os.path.join(BASE_DIR, ".venv-ocr", "Scripts", "python.exe")
MAIN = os.path.join(BASE_DIR, "main.py")


def _log_handle():
    """打开日志；打不开也不影响服务启动，只在窗口里提示一句。"""
    try:
        handle = open(LOG_PATH, "w", encoding="utf-8", buffering=1)
        handle.write("===== 启动 %s =====\n" % time.strftime("%Y-%m-%d %H:%M:%S"))
        return handle
    except Exception as exc:  # noqa: BLE001
        print("提示：无法写入日志文件（%s），服务仍会继续启动。" % exc)
        return None


def main() -> int:
    # 控制台编码在不同机器上不一致，这里允许替换无法编码的字符，避免因为一个
    # 打印语句把启动流程打断。
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:  # noqa: BLE001
            pass

    print("=" * 54)
    print("  办公文档转换器 - 启动服务")
    print("  本机访问: http://127.0.0.1:8000")
    print("  按 Ctrl+C 停止服务；日志同时写入 _start_log.txt")
    print("=" * 54)
    print()

    if not os.path.isfile(PYTHON):
        print("错误：找不到虚拟环境解释器 %s" % PYTHON)
        print("请确认 .venv-ocr 目录完整，并把这条消息告诉 Codex。")
        return 1
    if not os.path.isfile(MAIN):
        print("错误：找不到 %s，请把这条消息告诉 Codex。" % MAIN)
        return 1

    log = _log_handle()
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"

    try:
        process = subprocess.Popen(
            [PYTHON, MAIN],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", env=env, cwd=BASE_DIR,
        )
    except Exception as exc:  # noqa: BLE001
        print("错误：启动服务失败：%s" % exc)
        if log:
            log.write("启动失败：%s\n" % exc)
            log.close()
        return 1

    try:
        hinted = False
        for line in process.stdout:
            sys.stdout.write(line)
            sys.stdout.flush()
            if log:
                log.write(line)
            if not hinted and "Uvicorn running on" in line:
                hinted = True
                print()
                print("*" * 56)
                print("  服务已就绪：请不要关闭本窗口！")
                print("  浏览器打开: http://127.0.0.1:8000")
                print("  关闭本窗口 = 停止服务（要停止时按 Ctrl+C）")
                print("*" * 56)
                print()
    except KeyboardInterrupt:
        # 用户按了 Ctrl+C：交给下面统一收尾并打印结果
        pass
    finally:
        try:
            process.wait(timeout=30)
        except Exception:  # noqa: BLE001
            process.terminate()
        if process.stdout:
            process.stdout.close()

    print()
    if process.returncode == 0:
        print("服务已停止。日志文件: _start_log.txt")
    else:
        print("服务异常退出，返回码 %s。请把 _start_log.txt 发给 Codex。" % process.returncode)
    if log:
        log.write("===== %s 服务已停止，返回码 %s =====\n"
                  % (time.strftime("%Y-%m-%d %H:%M:%S"), process.returncode))
        log.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
