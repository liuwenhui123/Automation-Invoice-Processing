from __future__ import annotations

from decimal import Decimal
from typing import Iterable

from .models import CategorySummary, InvoiceRecord


DEFAULT_CATEGORIES = ("个人垫付", "差旅", "对公转账")


def summarize_category_totals(
    records: Iterable[InvoiceRecord],
    categories: tuple[str, ...] = DEFAULT_CATEGORIES,
) -> list[CategorySummary]:
    summaries = [CategorySummary(name=category) for category in categories]
    summary_by_name = {summary.name: summary for summary in summaries}

    for record in records:
        amount = record.total_amount if record.total_amount is not None else Decimal("0")
        seen_categories: set[str] = set()
        for category in record.categories:
            if category in seen_categories:
                continue
            seen_categories.add(category)

            summary = summary_by_name.get(category)
            if summary is None:
                summary = CategorySummary(name=category)
                summary_by_name[category] = summary
                summaries.append(summary)

            summary.invoice_count += 1
            summary.total_amount += amount

    return summaries
