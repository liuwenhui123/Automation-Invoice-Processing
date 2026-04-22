from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from invoice_app import InvoiceRecord, iter_pdf_files, parse_invoice, resolve_attachment_metadata
from invoice_app.ui_improved import run_improved_ui


INVALID_FILENAME_CHARS = '<>:"/\\|?*'
DEFAULT_HEADERS = (
    "发票号码",
    "开票日期",
    "购买方名称",
    "购买方税号",
    "销售方名称",
    "销售方税号",
    "发票金额",
    "发票税额",
    "发票总额",
    "总额大写",
    "物品名称",
)


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    return text.replace(" ", " ").replace("　", " ").replace("￥", "¥")


def format_amount(amount: Decimal) -> str:
    return str(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def sanitize_filename_part(value: str) -> str:
    value = normalize_text(value).strip()
    value = re.sub(r"\s+", "", value)
    for char in INVALID_FILENAME_CHARS:
        value = value.replace(char, "_")
    return value.rstrip(" .")


def main() -> int:
    parser = argparse.ArgumentParser(description="发票批量处理工具")
    parser.add_argument(
        "--ui",
        action="store_true",
        help="启动图形界面",
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="强制使用命令行模式",
    )
    args = parser.parse_args()

    if args.ui or (not args.cli):
        run_improved_ui()
        return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
