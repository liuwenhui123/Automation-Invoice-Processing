from __future__ import annotations

import re
import sys
import unicodedata
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

from .models import InvoiceRecord


ATTACHMENT_TITLES = (
    "销售货物或者提供应税劳务、服务清单",
    "销售货物或提供应税劳务、服务清单",
    "增值税发票销货清单",
    "销货清单",
    "销售清单",
)
DATE_PATTERN = re.compile(r"\d{4}年\d{2}月\d{2}日")
TAX_CODE_PATTERN = re.compile(r"[0-9A-Z]{15,20}")


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    return text.replace("\u00A0", " ").replace("\u3000", " ").replace("￥", "¥")


def flatten_text(text: str) -> str:
    return re.sub(r"\s+", "", normalize_text(text))


def clean_line(line: str) -> str:
    return re.sub(r"\s+", " ", normalize_text(line)).strip()


def compact_line(line: str) -> str:
    return re.sub(r"\s+", "", normalize_text(line))


def compact_lines(text: str) -> list[str]:
    return [compact_line(line) for line in text.splitlines() if compact_line(line)]


def is_numeric_token(value: str) -> bool:
    return bool(re.fullmatch(r"-?\d+(?:\.\d+)?", value))


def is_tax_rate_token(value: str) -> bool:
    return bool(
        re.fullmatch(r"\d+%", value)
        or value in {"免税", "不征税", "出口零税率", "普通零税率"}
    )


def extract_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    pages: list[str] = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages)


def detect_attachment(text: str) -> bool:
    flat = flatten_text(text)
    head = flat[:200]
    return any(title in head for title in ATTACHMENT_TITLES)


def is_invoice_like_pdf(text: str, is_attachment: bool) -> bool:
    if is_attachment:
        return True

    flat = flatten_text(text)
    title_hits = (
        "发票" in flat,
        "电子发票" in flat,
        "增值税" in flat,
        "普通发票" in flat,
        "专用发票" in flat,
        "全电发票" in flat,
        "数电发票" in flat,
    )
    field_hits = (
        "发票号码" in flat,
        "开票日期" in flat,
        "价税合计" in flat,
        "销售方" in flat or "销方" in flat,
        "购买方" in flat or "购方" in flat or "买方" in flat,
    )
    return sum(title_hits) >= 1 and sum(field_hits) >= 2


def parse_invoice_number(text: str) -> str | None:
    lines = compact_lines(text)
    for index, line in enumerate(lines):
        if not re.fullmatch(r"\d{8,20}", line):
            continue
        previous = "".join(lines[max(0, index - 4) : index])
        next_line = lines[index + 1] if index + 1 < len(lines) else ""
        if "发票号码" in previous or DATE_PATTERN.fullmatch(next_line):
            return line

    candidates = (
        normalize_text(text),
        flatten_text(text),
    )
    patterns = (
        r"发票号码[:：]?\s*(\d{8,20})",
        r"号码[:：]?\s*(\d{8,20})",
        r"票号[:：]?\s*(\d{8,20})",
    )
    for candidate in candidates:
        for pattern in patterns:
            match = re.search(pattern, candidate)
            if match:
                return match.group(1)
    return None


def cleanup_name(name: str) -> str:
    name = normalize_text(name)
    name = re.sub(r"\s+", "", name)
    name = re.sub(r"(统一社会信用代码|纳税人识别号|项目名称|开户行及账号).*$", "", name)
    name = name.strip("：:;；，,。.")
    name = re.sub(r"^(名称[:：]?)", "", name)
    return name


def parse_invoice_date(text: str) -> str | None:
    match = DATE_PATTERN.search(normalize_text(text))
    if match:
        return match.group(0)
    return None


def is_party_info_noise(value: str, invoice_number: str | None) -> bool:
    compact = compact_line(value)
    if not compact:
        return True
    if invoice_number and compact == invoice_number:
        return True
    if DATE_PATTERN.fullmatch(compact):
        return True
    if any(
        token in compact
        for token in (
            "电子发票",
            "普通发票",
            "专用发票",
            "发票号码",
            "开票日期",
            "统一社会信用代码",
            "纳税人识别号",
            "购买方信息",
            "销售方信息",
            "销方信息",
            "购方信息",
            "项目名称",
            "价税合计",
            "合计",
            "备注",
            "开票人",
            "税务局",
        )
    ):
        return True
    return False


def parse_party_details(
    text: str,
    invoice_number: str | None,
) -> tuple[str | None, str | None, str | None, str | None]:
    lines = [clean_line(line) for line in text.splitlines() if clean_line(line)]
    compact = [compact_line(line) for line in lines]
    info_lines = compact

    code_indexes: list[int] = []
    for index, line in enumerate(info_lines):
        if not TAX_CODE_PATTERN.fullmatch(line):
            continue
        if invoice_number and line == invoice_number:
            continue
        if DATE_PATTERN.fullmatch(line):
            continue
        code_indexes.append(index)

    def find_name_before(code_index: int) -> str | None:
        for cursor in range(code_index - 1, -1, -1):
            candidate = cleanup_name(info_lines[cursor])
            if not candidate:
                continue
            if is_party_info_noise(candidate, invoice_number):
                continue
            if TAX_CODE_PATTERN.fullmatch(compact_line(candidate)):
                continue
            return candidate
        return None

    buyer_name = buyer_tax_code = seller_name = seller_tax_code = None
    if len(code_indexes) >= 1:
        buyer_tax_code = info_lines[code_indexes[0]]
        buyer_name = find_name_before(code_indexes[0])
    if len(code_indexes) >= 2:
        seller_tax_code = info_lines[code_indexes[1]]
        seller_name = find_name_before(code_indexes[1])

    if buyer_name and seller_name:
        return buyer_name, buyer_tax_code, seller_name, seller_tax_code

    flat = flatten_text(text)
    name_matches = re.findall(
        r"名称[:：](.+?)(?=统一社会信用代码|纳税人识别号|项目名称|价税合计|合计)",
        flat,
    )
    code_matches = re.findall(
        r"(?:统一社会信用代码/纳税人识别号|纳税人识别号)[:：]([0-9A-Z]{15,20})",
        flat,
    )
    cleaned_names = [cleanup_name(name) for name in name_matches if cleanup_name(name)]
    if not buyer_name and cleaned_names:
        buyer_name = cleaned_names[0]
    if not seller_name and len(cleaned_names) >= 2:
        seller_name = cleaned_names[1]
    if not buyer_tax_code and code_matches:
        buyer_tax_code = code_matches[0]
    if not seller_tax_code and len(code_matches) >= 2:
        seller_tax_code = code_matches[1]
    return buyer_name, buyer_tax_code, seller_name, seller_tax_code


def parse_total_amount(text: str) -> Decimal | None:
    lines = [clean_line(line) for line in text.splitlines() if clean_line(line)]
    for line in lines:
        if "¥" not in line:
            continue
        if any(marker in line for marker in ("圆", "元", "角", "分", "整")):
            match = re.search(r"¥\s*(\d+(?:\.\d{1,2})?)", line)
            if match:
                try:
                    return Decimal(match.group(1))
                except InvalidOperation:
                    pass

    candidates = (
        flatten_text(text),
        normalize_text(text),
    )
    patterns = (
        r"价税合计[(（]大写[)）].*?[(（]小写[)）].*?¥(\d+(?:\.\d{1,2})?)",
        r"小写[)）].*?¥(\d+(?:\.\d{1,2})?)",
    )
    for candidate in candidates:
        for pattern in patterns:
            match = re.search(pattern, candidate)
            if not match:
                continue
            try:
                return Decimal(match.group(1))
            except InvalidOperation:
                continue
    return None


def parse_amount_and_tax(text: str) -> tuple[Decimal | None, Decimal | None]:
    lines = [clean_line(line) for line in text.splitlines() if clean_line(line)]
    for line in lines:
        if "¥" not in line:
            continue
        if any(marker in line for marker in ("圆", "元", "角", "分", "整")):
            continue
        amounts = re.findall(r"(\d+(?:\.\d{1,2})?)", line)
        if len(amounts) >= 2:
            try:
                return Decimal(amounts[0]), Decimal(amounts[1])
            except InvalidOperation:
                continue

    flat = flatten_text(text)
    patterns = (
        r"合计¥?(\d+(?:\.\d{1,2})?)¥?(\d+(?:\.\d{1,2})?)",
        r"(\d+(?:\.\d{1,2})?)¥(\d+(?:\.\d{1,2})?)¥",
    )
    for pattern in patterns:
        match = re.search(pattern, flat)
        if not match:
            continue
        try:
            return Decimal(match.group(1)), Decimal(match.group(2))
        except InvalidOperation:
            continue
    return None, None


def parse_total_amount_string(text: str) -> str | None:
    lines = [clean_line(line) for line in text.splitlines() if clean_line(line)]
    for line in lines:
        match = re.search(r"([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+).*?¥", line)
        if match:
            return match.group(1)

    flat = flatten_text(text)
    match = re.search(r"([零壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元角分整]+)¥\d+(?:\.\d{1,2})?", flat)
    if match:
        return match.group(1)
    return None


def looks_like_item_fragment(line: str) -> bool:
    compact = compact_line(line)
    if not compact:
        return False
    if any(token in compact for token in ("合计", "价税合计", "备注", "备", "开票人")):
        return False
    if TAX_CODE_PATTERN.fullmatch(compact) or DATE_PATTERN.fullmatch(compact):
        return False
    if "¥" in compact:
        return False
    if "%" in compact:
        return False
    if len(re.findall(r"\d+(?:\.\d+)?", compact)) >= 2:
        return False
    return bool(re.search(r"[\u4e00-\u9fff\*]", compact))


def extract_name_from_detail_line(line: str) -> str | None:
    tokens = line.split()
    if not tokens:
        return None
    rate_index = next(
        (index for index, token in enumerate(tokens) if is_tax_rate_token(token)),
        None,
    )
    if rate_index is None:
        return None

    strip_count = 3
    if (
        rate_index >= 4
        and not is_numeric_token(tokens[rate_index - 4])
        and len(tokens[rate_index - 4]) <= 3
    ):
        strip_count = 4
    name_tokens = tokens[: max(0, rate_index - strip_count)]
    name = "".join(name_tokens)
    name = re.sub(r"^[*]+|[*]+$", "*", name)
    return name or None


def is_detail_row(line: str) -> bool:
    tokens = line.split()
    if not tokens:
        return False
    if not any(is_tax_rate_token(token) for token in tokens):
        return False
    numeric_count = sum(1 for token in tokens if is_numeric_token(token))
    return numeric_count >= 2


def parse_item_names(text: str) -> list[str]:
    lines = [clean_line(line) for line in text.splitlines() if clean_line(line)]
    items: list[str] = []
    buffer: list[str] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        compact = compact_line(line)

        if is_detail_row(line):
            name = extract_name_from_detail_line(line) or "".join(buffer)
            buffer.clear()

            next_index = index + 1
            while next_index < len(lines):
                next_line = lines[next_index]
                next_compact = compact_line(next_line)
                if any(token in next_compact for token in ("合计", "价税合计", "备注", "备", "开票人")):
                    break
                if is_detail_row(next_line):
                    break
                if looks_like_item_fragment(next_line):
                    name += next_compact
                    next_index += 1
                    continue
                break

            cleaned = cleanup_name(name).replace(" ", "")
            if cleaned:
                items.append(cleaned)
            index = next_index
            continue

        if "*" in compact or buffer:
            if looks_like_item_fragment(line):
                buffer.append(compact)
            else:
                buffer.clear()
        elif looks_like_item_fragment(line) and "*" in compact:
            buffer.append(compact)
        index += 1

    if buffer:
        cleaned = cleanup_name("".join(buffer)).replace(" ", "")
        if cleaned:
            items.append(cleaned)
    return items


def parse_invoice(path: Path) -> InvoiceRecord:
    text = extract_pdf_text(path)
    is_attachment = detect_attachment(text)
    is_invoice_like = is_invoice_like_pdf(text, is_attachment)
    invoice_number = parse_invoice_number(text)
    invoice_date = parse_invoice_date(text)
    buyer_name, buyer_tax_code, seller_name, seller_tax_code = parse_party_details(
        text, invoice_number
    )
    amount, tax_amount = parse_amount_and_tax(text)
    total_amount = parse_total_amount(text)
    total_amount_string = parse_total_amount_string(text)
    record = InvoiceRecord(
        source=path,
        is_invoice_like=is_invoice_like,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        buyer_name=buyer_name,
        buyer_tax_code=buyer_tax_code,
        seller_name=seller_name,
        seller_tax_code=seller_tax_code,
        amount=amount,
        tax_amount=tax_amount,
        total_amount=total_amount,
        total_amount_string=total_amount_string,
        item_names=parse_item_names(text),
        is_attachment=is_attachment,
    )
    if not record.is_invoice_like:
        record.warnings.append("非发票类 PDF，已跳过")
        return record
    if not record.invoice_number:
        record.warnings.append("未识别到发票号码")
    if not record.seller_name:
        record.warnings.append("未识别到销售方名称")
    if not record.buyer_name:
        record.warnings.append("未识别到购买方名称")
    if not record.invoice_date:
        record.warnings.append("未识别到开票日期")
    if record.total_amount is None:
        record.warnings.append("未识别到价税合计金额")
    return record


def iter_pdf_files(inputs: Iterable[str], recursive: bool) -> list[Path]:
    paths = list(inputs) or ["."]
    files: list[Path] = []
    seen: set[str] = set()
    for raw in paths:
        path = Path(raw).expanduser().resolve()
        if path.is_file():
            candidates = [path] if path.suffix.lower() == ".pdf" else []
        elif path.is_dir():
            iterator = path.rglob("*.pdf") if recursive else path.glob("*.pdf")
            candidates = [item.resolve() for item in iterator if item.is_file()]
        else:
            print(f"[WARN] 路径不存在，已跳过: {path}", file=sys.stderr)
            continue

        for candidate in sorted(candidates, key=lambda item: str(item).lower()):
            key = str(candidate).casefold()
            if key in seen:
                continue
            seen.add(key)
            files.append(candidate)
    return files


def resolve_attachment_metadata(records: list[InvoiceRecord]) -> None:
    main_records = {
        record.invoice_number: record
        for record in records
        if record.invoice_number and not record.is_attachment
    }
    for record in records:
        if not record.is_attachment or not record.invoice_number:
            continue
        main_record = main_records.get(record.invoice_number)
        if not main_record:
            continue
        if not record.invoice_date:
            record.invoice_date = main_record.invoice_date
        if not record.buyer_name:
            record.buyer_name = main_record.buyer_name
        if not record.buyer_tax_code:
            record.buyer_tax_code = main_record.buyer_tax_code
        if not record.seller_name:
            record.seller_name = main_record.seller_name
        if not record.seller_tax_code:
            record.seller_tax_code = main_record.seller_tax_code
        if record.amount is None:
            record.amount = main_record.amount
        if record.tax_amount is None:
            record.tax_amount = main_record.tax_amount
        if record.total_amount is None:
            record.total_amount = main_record.total_amount
        if not record.total_amount_string:
            record.total_amount_string = main_record.total_amount_string
