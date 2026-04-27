from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook

from .classification import DEFAULT_CATEGORIES, summarize_category_totals
from .models import InvoiceRecord


DETAIL_HEADERS = (
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
SUMMARY_HEADERS = ("类别名称", "发票数量", "汇总金额")


def _as_excel_number(value: Decimal | None) -> float | None:
    if value is None:
        return None
    return float(value)


def export_records_with_summary(
    records: Iterable[InvoiceRecord],
    output_path: Path,
    *,
    category_names: tuple[str, ...] = DEFAULT_CATEGORIES,
) -> None:
    records = list(records)
    workbook = Workbook()

    detail_sheet = workbook.active
    detail_sheet.title = "发票明细"
    detail_sheet.append(list(DETAIL_HEADERS))

    for record in records:
        detail_sheet.append(
            [
                record.invoice_number or "",
                record.invoice_date or "",
                record.buyer_name or "",
                record.buyer_tax_code or "",
                record.seller_name or "",
                record.seller_tax_code or "",
                _as_excel_number(record.amount),
                _as_excel_number(record.tax_amount),
                _as_excel_number(record.total_amount),
                record.total_amount_string or "",
                "、".join(record.item_names) if record.item_names else "",
            ]
        )

    summary_sheet = workbook.create_sheet("分类汇总")
    summary_sheet.append(list(SUMMARY_HEADERS))

    for summary in summarize_category_totals(records, category_names):
        if summary.invoice_count <= 0:
            continue
        summary_sheet.append(
            [
                summary.name,
                summary.invoice_count,
                _as_excel_number(summary.total_amount),
            ]
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
