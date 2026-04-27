from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from shutil import copy2

from .archive import plan_archive_targets
from .classification import DEFAULT_CATEGORIES, summarize_category_totals
from .export_excel import export_records_with_summary
from .models import CategorySummary, InvoiceRecord
from .naming import format_amount, sanitize_path_component


@dataclass
class ProcessOptions:
    output_root: Path
    category_names: tuple[str, ...] = DEFAULT_CATEGORIES
    excel_output: Path | None = None


@dataclass
class PreviewResult:
    records: list[InvoiceRecord]
    category_summaries: list[CategorySummary]
    review_notes: list[str] = field(default_factory=list)


@dataclass
class ProcessResult:
    success_count: int = 0
    failure_count: int = 0
    skipped_count: int = 0
    failures: list[str] = field(default_factory=list)
    generated_folders: list[Path] = field(default_factory=list)
    excel_output_path: Path | None = None
    records: list[InvoiceRecord] = field(default_factory=list)


def add_category(categories: list[str], name: str) -> list[str]:
    category_name = name.strip()
    if not category_name or category_name in categories:
        return list(categories)
    return [*categories, category_name]


def rename_category(
    categories: list[str],
    old_name: str,
    new_name: str,
    records: list[InvoiceRecord],
) -> list[str]:
    previous = old_name.strip()
    target = new_name.strip()
    if not previous or not target:
        return list(categories)
    if previous not in categories:
        return list(categories)

    renamed_categories = [target if item == previous else item for item in categories]
    deduped_categories: list[str] = []
    seen: set[str] = set()
    for item in renamed_categories:
        if item in seen:
            continue
        seen.add(item)
        deduped_categories.append(item)

    for record in records:
        record.categories = _dedupe_categories(
            [target if item == previous else item for item in record.categories]
        )
    return deduped_categories


def remove_category(
    categories: list[str],
    name: str,
    records: list[InvoiceRecord],
) -> list[str]:
    target = name.strip()
    remaining = [item for item in categories if item != target]
    for record in records:
        record.categories = [item for item in record.categories if item != target]
    return remaining


def apply_category_to_records(
    records: list[InvoiceRecord],
    category_name: str,
    *,
    target_indices: list[int] | None = None,
    replace: bool = True,
) -> list[InvoiceRecord]:
    category = category_name.strip()
    if not category:
        return records
    indices = target_indices if target_indices is not None else list(range(len(records)))
    for index in indices:
        if index < 0 or index >= len(records):
            continue
        record = records[index]
        if replace:
            record.categories = [category]
        else:
            record.categories = _dedupe_categories([*record.categories, category])
    return records


def clear_categories(
    records: list[InvoiceRecord],
    *,
    target_indices: list[int] | None = None,
) -> list[InvoiceRecord]:
    indices = target_indices if target_indices is not None else list(range(len(records)))
    for index in indices:
        if index < 0 or index >= len(records):
            continue
        records[index].categories = []
    return records


def _dedupe_categories(categories: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for item in categories:
        if not item or item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


def _effective_categories(record: InvoiceRecord) -> list[str]:
    return _dedupe_categories(record.categories)


def _build_main_stem(record: InvoiceRecord) -> str:
    suffix = record.invoice_suffix or (
        sanitize_path_component(record.source.stem) if record.source is not None else "invoice"
    )
    seller = sanitize_path_component(record.seller_name or "未知商家")
    amount = format_amount(record.total_amount or Decimal("0"))
    parts = [suffix]
    if record.display_name.strip():
        parts.append(sanitize_path_component(record.display_name))
    parts.extend([seller, amount])
    if record.is_attachment:
        parts.append("附件")
    return "_".join(parts)


def _ensure_unique_path(path: Path, reserved: set[str]) -> Path:
    candidate = path
    index = 2
    while str(candidate).casefold() in reserved or candidate.exists():
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        index += 1
    reserved.add(str(candidate).casefold())
    return candidate


def _inherit_attachment_categories(records: list[InvoiceRecord]) -> None:
    category_map: dict[str, list[str]] = {}
    for record in records:
        if record.is_attachment:
            continue
        if not record.invoice_number:
            continue
        category_map[record.invoice_number] = _effective_categories(record)
    for record in records:
        if not record.is_attachment:
            continue
        if record.categories:
            continue
        if not record.invoice_number:
            continue
        inherited = category_map.get(record.invoice_number)
        if inherited:
            record.categories = list(inherited)


def preview_records(records: list[InvoiceRecord], options: ProcessOptions) -> PreviewResult:
    review_notes: list[str] = []
    previewed_records: list[InvoiceRecord] = []
    category_names = options.category_names

    for record in records:
        problems: list[str] = []
        if not record.is_invoice_like:
            status = "已跳过"
            previewed_record = deepcopy(record)
            previewed_record.status = status
            previewed_records.append(previewed_record)
            source_name = record.source.name if record.source is not None else "未知来源"
            review_notes.append(f"{source_name}: 非发票类 PDF")
            continue
        if not record.categories:
            problems.append("未分类")
        if record.total_amount is None:
            problems.append("未识别到价税合计金额")

        status = "已就绪"
        if problems:
            status = "需人工确认"
            source_name = record.source.name if record.source is not None else "未知来源"
            review_notes.append(f"{source_name}: {'；'.join(problems)}")
        previewed_record = deepcopy(record)
        previewed_record.status = status
        previewed_records.append(previewed_record)

    for record in previewed_records:
        record.categories = _effective_categories(record)

    category_summaries = summarize_category_totals(previewed_records, category_names)
    return PreviewResult(
        records=previewed_records,
        category_summaries=category_summaries,
        review_notes=review_notes,
    )


def execute_records(records: list[InvoiceRecord], options: ProcessOptions) -> ProcessResult:
    working_records = deepcopy(records)
    _inherit_attachment_categories(working_records)

    result = ProcessResult()
    reserved_paths: set[str] = set()
    generated_folders: set[Path] = set()
    ready_records = [
        record
        for record in working_records
        if record.status == "已就绪" and record.is_invoice_like
    ]
    category_summaries = summarize_category_totals(ready_records, options.category_names)
    category_totals = {
        item.name: item.total_amount
        for item in category_summaries
        if item.invoice_count > 0
    }

    for index, record in enumerate(working_records):
        source_name = record.source.name if record.source is not None else "未知来源"
        record.categories = _effective_categories(record)
        if record.status != "已就绪" or not record.is_invoice_like:
            result.skipped_count += 1
            continue
        if record.source is None or not record.source.exists():
            result.failure_count += 1
            record.status = "失败"
            result.failures.append(f"{source_name}: 源文件不存在")
            continue

        try:
            record_root = (
                record.source.parent if record.source is not None else options.output_root
            )
            record_root.mkdir(parents=True, exist_ok=True)
            record.canonical_name = _build_main_stem(record)
            plan = plan_archive_targets(record, record_root, category_totals)

            if not record.categories:
                canonical_target = _ensure_unique_path(
                    record_root / Path(plan.canonical_path).name,
                    reserved_paths,
                )
                canonical_target.parent.mkdir(parents=True, exist_ok=True)
                record.source.rename(canonical_target)
                record.canonical_path = canonical_target
                record.archive_paths = []
            else:
                first_target = _ensure_unique_path(plan.copy_targets[0], reserved_paths)
                first_target.parent.mkdir(parents=True, exist_ok=True)
                record.source.rename(first_target)
                generated_folders.add(first_target.parent)

                archive_paths: list[Path] = [first_target]
                for copy_target in plan.copy_targets[1:]:
                    target_path = _ensure_unique_path(copy_target, reserved_paths)
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    copy2(first_target, target_path)
                    archive_paths.append(target_path)
                    generated_folders.add(target_path.parent)

                record.canonical_path = first_target
                record.archive_paths = archive_paths

            record.status = "已执行"
            result.success_count += 1
        except Exception as exc:  # noqa: BLE001
            result.failure_count += 1
            record.status = "失败"
            result.failures.append(f"{source_name}: {exc}")

        records[index] = record

    if options.excel_output is not None:
        export_records_with_summary(
            records,
            options.excel_output,
            category_names=options.category_names,
        )
        result.excel_output_path = options.excel_output

    result.generated_folders = sorted(generated_folders)
    result.records = working_records
    return result
