"""Office COM 隔离进程入口；由 converter.py 调用，不直接对外提供服务。"""
import json
import sys

import converter


def main() -> int:
    if len(sys.argv) not in (5, 6):
        raise ValueError("Office worker 参数错误")

    operation, src, dst, app_name = sys.argv[1:5]
    if operation == "pdf":
        converter._office_to_pdf(src, dst, app_name)
    elif operation == "probe":
        # 健康探测：空文档转 PDF，用于判断 Office 自动化当前是否可用
        converter._office_probe(dst, app_name)
    elif operation == "scan-docx":
        # 扫描件可编辑图层：OCR 结果通过 JSON 文件传入（argv[5]）
        with open(sys.argv[5], "r", encoding="utf-8") as fh:
            converter._save_scanned_editable_docx(src, dst, json.load(fh))
    elif operation == "format" and len(sys.argv) == 6:
        converter._office_convert_format(src, dst, app_name, int(sys.argv[5]))
    else:
        raise ValueError(f"未知 Office 操作: {operation}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
