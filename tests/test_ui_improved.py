from decimal import Decimal
from pathlib import Path

from invoice_app.models import InvoiceRecord
from invoice_app.ui_improved import ImprovedInvoiceApp


def test_refresh_preview_keeps_live_record_references_when_updating_summary():
    app = ImprovedInvoiceApp.__new__(ImprovedInvoiceApp)
    app.categories = ["个人垫付", "差旅", "对公转账"]
    app.workspace_root = Path("out")
    app.records = [
        InvoiceRecord(
            source=Path("a.pdf"),
            is_invoice_like=True,
            invoice_number="12345678",
            seller_name="测试商家A",
            total_amount=Decimal("1480.00"),
            categories=["个人垫付"],
        ),
        InvoiceRecord(
            source=Path("b.pdf"),
            is_invoice_like=True,
            invoice_number="87654321",
            seller_name="测试商家B",
            total_amount=Decimal("100.00"),
            categories=["个人垫付"],
        ),
    ]
    original_records = app.records
    captured = {}

    class BusyStub:
        def get(self):
            return False

    app.busy = BusyStub()
    app._refresh_summary = lambda summaries: captured.setdefault("summaries", summaries)

    ImprovedInvoiceApp.refresh_preview(app)

    assert app.records is original_records
    assert app.records[0] is original_records[0]
    assert app.records[1] is original_records[1]
    assert captured["summaries"][0].name == "个人垫付"
    assert captured["summaries"][0].invoice_count == 2


def test_refresh_preview_preserves_row_bound_record_objects_across_multiple_changes():
    app = ImprovedInvoiceApp.__new__(ImprovedInvoiceApp)
    app.categories = ["个人垫付", "差旅", "对公转账"]
    app.workspace_root = Path("out")
    app.records = [
        InvoiceRecord(
            source=Path("a.pdf"),
            is_invoice_like=True,
            invoice_number="12345678",
            seller_name="测试商家A",
            total_amount=Decimal("1480.00"),
            categories=[],
        ),
        InvoiceRecord(
            source=Path("b.pdf"),
            is_invoice_like=True,
            invoice_number="87654321",
            seller_name="测试商家B",
            total_amount=Decimal("100.00"),
            categories=[],
        ),
    ]
    first_row_record = app.records[0]
    second_row_record = app.records[1]
    summaries_seen = []

    class BusyStub:
        def get(self):
            return False

    app.busy = BusyStub()
    app._refresh_summary = lambda summaries: summaries_seen.append(summaries)

    first_row_record.categories = ["个人垫付"]
    ImprovedInvoiceApp.refresh_preview(app)

    second_row_record.categories = ["个人垫付"]
    ImprovedInvoiceApp.refresh_preview(app)

    assert app.records[0] is first_row_record
    assert app.records[1] is second_row_record
    assert app.records[0].categories == ["个人垫付"]
    assert app.records[1].categories == ["个人垫付"]
    assert summaries_seen[-1][0].name == "个人垫付"
    assert summaries_seen[-1][0].invoice_count == 2
