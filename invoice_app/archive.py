from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Mapping

from .models import ArchivePlan, InvoiceRecord
from .naming import build_category_folder_name, sanitize_path_component


def _resolve_main_name(record: InvoiceRecord) -> str:
    if record.canonical_name:
        return sanitize_path_component(record.canonical_name)
    if record.display_name:
        return sanitize_path_component(record.display_name)
    if record.source is not None:
        return sanitize_path_component(record.source.stem)
    return "invoice"


def plan_archive_targets(
    record: InvoiceRecord,
    archive_root: Path,
    category_totals: Mapping[str, Decimal],
) -> ArchivePlan:
    main_name = _resolve_main_name(record)
    suffix = record.source.suffix if record.source is not None and record.source.suffix else ".pdf"
    canonical_path = archive_root / f"{main_name}{suffix}"

    copy_targets: list[Path] = []
    seen_categories: set[str] = set()
    for category in record.categories:
        if category in seen_categories:
            continue
        seen_categories.add(category)
        category_total = category_totals.get(category)
        if category_total is None:
            continue
        folder_name = build_category_folder_name(category, category_total)
        copy_targets.append(archive_root / folder_name / canonical_path.name)

    return ArchivePlan(canonical_path=canonical_path, copy_targets=copy_targets)
