from __future__ import annotations

from decimal import Decimal


WINDOWS_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def sanitize_path_component(value: str) -> str:
    cleaned = value.strip()
    for character in WINDOWS_INVALID_FILENAME_CHARS:
        cleaned = cleaned.replace(character, "_")
    return cleaned.rstrip(" .") or "_"


def format_amount(amount: Decimal) -> str:
    return format(amount.quantize(Decimal("0.00")), "f")


def build_category_folder_name(category: str, amount: Decimal) -> str:
    return f"{sanitize_path_component(category)}_{format_amount(amount)}"
