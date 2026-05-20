from decimal import Decimal
from pathlib import Path

from invoice_app.models import InvoiceRecord
from invoice_app.service import ProcessOptions, execute_records, preview_records


def test_unclassified_invoice_is_renamed_without_category_folder(tmp_path):
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    source = source_dir / "raw.pdf"
    source.write_bytes(b"invoice")

    records = [
        InvoiceRecord(
            source=source,
            invoice_number="123456737345",
            seller_name="阿里云计算有限公司",
            total_amount=Decimal("252.00"),
            categories=[],
            status="已就绪",
        )
    ]

    result = execute_records(
        records,
        ProcessOptions(
            output_root=tmp_path / "output",
            category_names=("个人垫付", "差旅", "对公转账"),
        ),
    )

    renamed_files = list(source_dir.glob("*.pdf"))
    category_dirs = [path for path in source_dir.iterdir() if path.is_dir()]

    assert result.success_count == 1
    assert result.skipped_count == 0
    assert len(renamed_files) == 1
    assert renamed_files[0].name == "737345_阿里云计算有限公司_252.00.pdf"
    assert category_dirs == []
    assert result.records[0].canonical_path == renamed_files[0]
    assert result.records[0].archive_paths == []


def test_preview_allows_unclassified_invoice_with_amount_to_execute():
    records = [
        InvoiceRecord(
            source=Path("raw.pdf"),
            invoice_number="123456737345",
            seller_name="阿里云计算有限公司",
            total_amount=Decimal("252.00"),
            categories=[],
        )
    ]

    preview = preview_records(records, ProcessOptions(output_root=Path("out")))

    assert preview.records[0].status == "已就绪"
    assert preview.review_notes == []
