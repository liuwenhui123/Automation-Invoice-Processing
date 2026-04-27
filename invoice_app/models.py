from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path


@dataclass
class InvoiceRecord:
    source: Path | None = None
    is_invoice_like: bool = True
    invoice_number: str | None = None
    invoice_date: str | None = None
    buyer_name: str | None = None
    buyer_tax_code: str | None = None
    seller_name: str | None = None
    seller_tax_code: str | None = None
    amount: Decimal | None = None
    tax_amount: Decimal | None = None
    total_amount: Decimal | None = None
    total_amount_string: str | None = None
    item_names: list[str] = field(default_factory=list)
    is_attachment: bool = False
    target: Path | None = None
    warnings: list[str] = field(default_factory=list)
    display_name: str = ""
    categories: list[str] = field(default_factory=list)
    canonical_name: str | None = None
    canonical_path: Path | None = None
    archive_paths: list[Path] = field(default_factory=list)
    status: str = "待处理"

    @property
    def invoice_suffix(self) -> str | None:
        if not self.invoice_number:
            return None
        return self.invoice_number[-6:]


@dataclass
class CategorySummary:
    name: str
    invoice_count: int = 0
    total_amount: Decimal = Decimal("0")


@dataclass
class ArchivePlan:
    canonical_path: Path
    copy_targets: list[Path] = field(default_factory=list)
