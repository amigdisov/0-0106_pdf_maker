#!/usr/bin/env python3
"""Генератор PDF-накладных из CSV/JSON и HTML-шаблонов."""

from __future__ import annotations

import html
import json
import math
import os
import platform
import re
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
FONTS_DIR = BASE_DIR / "fonts"

VAT_RATE = 0.20
PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")
ITEM_ROW_RE = re.compile(
    r"<tr\b[^>]*>.*?\{\{\s*product\s*\}\}.*?</tr>",
    re.IGNORECASE | re.DOTALL,
)

ID_KEYS = ("invoice_id", "invoiceid", "invoice", "id", "номер")
DATE_KEYS = ("date", "invoice_date", "дата")
PRODUCT_KEYS = ("product", "name", "товар", "наименование")
PRICE_KEYS = ("price", "цена")
QUANTITY_KEYS = ("quantity", "qty", "количество", "кол-во", "кол_во")


@dataclass
class Item:
    product: str
    price: float
    quantity: float

    @property
    def line_sum(self) -> float:
        return self.price * self.quantity


@dataclass
class Invoice:
    invoice_id: str
    items: list[Item] = field(default_factory=list)
    date: str = ""

    @property
    def total(self) -> float:
        return sum(item.line_sum for item in self.items)

    @property
    def vat(self) -> float:
        """НДС, уже включённый в сумму (ставка 20%)."""
        return self.total * VAT_RATE / (1 + VAT_RATE)


def configure_stdio() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8")
        except Exception:
            pass


def prepare_weasyprint_env() -> None:
    """На Windows подсказывает WeasyPrint, где лежат DLL Pango."""
    if platform.system() != "Windows":
        return
    if os.environ.get("WEASYPRINT_DLL_DIRECTORIES"):
        return
    program_files = Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
    candidates = [
        Path(r"C:\msys64\ucrt64\bin"),
        Path(r"C:\msys64\mingw64\bin"),
        Path.home() / "msys64" / "ucrt64" / "bin",
        program_files / "GTK3-Runtime Win64" / "bin",
    ]
    existing = [str(path) for path in candidates if path.is_dir()]
    if existing:
        os.environ["WEASYPRINT_DLL_DIRECTORIES"] = os.pathsep.join(existing)


def as_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and (math.isnan(value) or value.is_integer()):
        if math.isnan(value):
            return ""
        return str(int(value))
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def parse_number(value: object) -> float | None:
    text = as_text(value).replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def format_date(value: object) -> str:
    text = as_text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:10], fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def format_money(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ")


def format_quantity(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return format_money(value)


def normalize_key(key: object) -> str:
    return str(key).strip().lower().replace(" ", "_").replace("-", "_")


def first_present(source: dict, keys: tuple[str, ...]) -> object:
    normalized = {normalize_key(key): value for key, value in source.items()}
    for key in keys:
        if key in normalized and as_text(normalized[key]):
            return normalized[key]
    return None


def item_from_mapping(source: dict) -> Item | None:
    product = as_text(first_present(source, PRODUCT_KEYS))
    price = parse_number(first_present(source, PRICE_KEYS))
    quantity = parse_number(first_present(source, QUANTITY_KEYS))
    if not product or price is None or quantity is None:
        return None
    return Item(product=product, price=price, quantity=quantity)


def invoice_from_mapping(source: dict, fallback_id: str) -> Invoice:
    invoice_id = as_text(first_present(source, ID_KEYS)) or fallback_id
    date = format_date(first_present(source, DATE_KEYS))
    raw_items = source.get("items") or source.get("products") or source.get("lines") or []
    items: list[Item] = []
    if isinstance(raw_items, list):
        for row in raw_items:
            if isinstance(row, dict):
                item = item_from_mapping(row)
                if item is not None:
                    items.append(item)
    return Invoice(invoice_id=invoice_id, items=items, date=date)


def invoices_from_rows(rows: list[dict], fallback_id: str) -> list[Invoice]:
    grouped: dict[str, Invoice] = {}
    order: list[str] = []
    for row in rows:
        invoice_id = as_text(first_present(row, ID_KEYS)) or fallback_id
        if invoice_id not in grouped:
            grouped[invoice_id] = Invoice(
                invoice_id=invoice_id,
                date=format_date(first_present(row, DATE_KEYS)),
            )
            order.append(invoice_id)
        invoice = grouped[invoice_id]
        if not invoice.date:
            invoice.date = format_date(first_present(row, DATE_KEYS))
        item = item_from_mapping(row)
        if item is not None:
            invoice.items.append(item)
    return [grouped[invoice_id] for invoice_id in order]


def load_csv(path: Path) -> list[Invoice]:
    frame = pd.read_csv(path, encoding="utf-8-sig")
    frame.columns = [normalize_key(column) for column in frame.columns]
    rows = frame.to_dict(orient="records")
    return invoices_from_rows(rows, fallback_id=path.stem)


def load_json(path: Path) -> list[Invoice]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    fallback_id = path.stem

    if isinstance(payload, dict):
        invoices = payload.get("invoices")
        if isinstance(invoices, list):
            payload = invoices
        elif payload and all(isinstance(value, dict) for value in payload.values()):
            documents = []
            for key, value in payload.items():
                document = dict(value)
                document.setdefault("invoice_id", key)
                documents.append(document)
            payload = documents
        else:
            payload = [payload]

    if not isinstance(payload, list):
        raise ValueError(f"Неподдерживаемая структура JSON: {path.name}")

    if not payload:
        return []

    if all(isinstance(item, dict) and not any(key in item for key in ("items", "products", "lines")) for item in payload):
        if any(first_present(item, PRODUCT_KEYS) is not None for item in payload):
            return invoices_from_rows(payload, fallback_id=fallback_id)

    invoices: list[Invoice] = []
    for index, item in enumerate(payload, start=1):
        if not isinstance(item, dict):
            continue
        invoices.append(invoice_from_mapping(item, fallback_id=f"{fallback_id}-{index}"))
    return invoices


def load_invoices(path: Path) -> list[Invoice]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return load_csv(path)
    if suffix == ".json":
        return load_json(path)
    raise ValueError(f"Неподдерживаемый формат: {path.name}")


def list_files(directory: Path, suffixes: tuple[str, ...]) -> list[Path]:
    if not directory.is_dir():
        return []
    found: dict[Path, Path] = {}
    for path in directory.iterdir():
        if path.is_file() and path.suffix.lower() in suffixes and not path.name.startswith("."):
            found[path.resolve()] = path
    return sorted(found.values(), key=lambda item: item.name.lower())


def fill_placeholders(template: str, values: dict[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        return values.get(match.group(1), match.group(0))

    return PLACEHOLDER_RE.sub(replace, template)


def apply_fonts(template_html: str) -> str:
    """Подключает DejaVu Sans абсолютным путём, чтобы кириллица была в PDF."""
    regular = (FONTS_DIR / "DejaVuSans.ttf").resolve().as_uri()
    bold = (FONTS_DIR / "DejaVuSans-Bold.ttf").resolve().as_uri()
    template_html = template_html.replace("../fonts/DejaVuSans-Bold.ttf", bold)
    template_html = template_html.replace("../fonts/DejaVuSans.ttf", regular)
    if "@font-face" in template_html:
        return template_html

    face = (
        "@font-face {"
        f' font-family: "DejaVu Sans"; src: url("{regular}");'
        " font-weight: 400; font-style: normal; }"
        "@font-face {"
        f' font-family: "DejaVu Sans"; src: url("{bold}");'
        " font-weight: 700; font-style: normal; }"
    )
    if "<style>" in template_html:
        return template_html.replace("<style>", "<style>" + face, 1)
    return face + template_html


def render_html(template_html: str, invoice: Invoice) -> str:
    template_html = apply_fonts(template_html)
    match = ITEM_ROW_RE.search(template_html)
    if match is None:
        raise ValueError("В шаблоне нет строки таблицы с {{ product }}.")

    row_template = match.group(0)
    rendered_rows: list[str] = []
    for index, item in enumerate(invoice.items, start=1):
        rendered_rows.append(
            fill_placeholders(
                row_template,
                {
                    "n": str(index),
                    "product": html.escape(item.product),
                    "price": format_money(item.price),
                    "quantity": format_quantity(item.quantity),
                    "sum": format_money(item.line_sum),
                },
            )
        )

    document = template_html[: match.start()] + "\n".join(rendered_rows) + template_html[match.end() :]
    total = format_money(invoice.total)
    return fill_placeholders(
        document,
        {
            "invoice_id": html.escape(invoice.invoice_id),
            "date": html.escape(invoice.date or "—"),
            "total": total,
            "vat": format_money(invoice.vat),
            "items_count": str(len(invoice.items)),
        },
    )


def safe_filename(invoice_id: str) -> str:
    cleaned = re.sub(r"[^\w.\-]+", "_", invoice_id, flags=re.UNICODE).strip("._")
    return cleaned or "invoice"


def write_pdf(html_document: str, template_path: Path, output_path: Path) -> None:
    prepare_weasyprint_env()
    try:
        from weasyprint import HTML
    except OSError as error:
        raise RuntimeError(weasyprint_help(error)) from error

    output_path.parent.mkdir(parents=True, exist_ok=True)
    HTML(string=html_document, base_url=str(template_path.parent)).write_pdf(output_path)


def weasyprint_help(error: BaseException) -> str:
    lines = [
        "Не удалось загрузить WeasyPrint. Для PDF нужны библиотеки Pango.",
        f"Подробности: {error}",
    ]
    if platform.system() == "Windows":
        lines.extend(
            [
                "Windows: установите MSYS2 (https://www.msys2.org/) и в оболочке UCRT64 выполните:",
                "  pacman -S mingw-w64-ucrt-x86_64-pango",
                "Затем перед запуском задайте переменную:",
                r"  set WEASYPRINT_DLL_DIRECTORIES=C:\msys64\ucrt64\bin",
            ]
        )
    elif platform.system() == "Darwin":
        lines.append("macOS: brew install pango")
    return "\n".join(lines)


def open_pdf(path: Path) -> None:
    system = platform.system()
    if system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
        return
    if system == "Darwin":
        subprocess.run(["open", str(path)], check=False)
        return
    subprocess.run(["xdg-open", str(path)], check=False)


def print_banner() -> None:
    print()
    print("=" * 48)
    print("  Генератор PDF-накладных")
    print("=" * 48)


def print_options(title: str, labels: list[str]) -> None:
    print()
    print(title)
    if not labels:
        print("  (ничего не найдено)")
        return
    for index, label in enumerate(labels, start=1):
        print(f"  {index}. {label}")


def prompt_choice(label: str, count: int) -> int:
    while True:
        raw = input(f"\n{label}: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= count:
            return int(raw) - 1
        print(f"Введите номер от 1 до {count}.")


def invoice_label(invoice: Invoice) -> str:
    details = [f"{len(invoice.items)} поз."]
    if invoice.date:
        details.append(invoice.date)
    return f"{invoice.invoice_id}  ({', '.join(details)})"


def check_weasyprint() -> None:
    prepare_weasyprint_env()
    try:
        from weasyprint import HTML  # noqa: F401
    except OSError as error:
        raise RuntimeError(weasyprint_help(error)) from error


def ensure_fonts() -> None:
    regular = FONTS_DIR / "DejaVuSans.ttf"
    bold = FONTS_DIR / "DejaVuSans-Bold.ttf"
    if regular.is_file() and bold.is_file():
        return
    missing = [path.name for path in (regular, bold) if not path.is_file()]
    raise FileNotFoundError(
        "Не найдены шрифты DejaVu Sans: "
        + ", ".join(missing)
        + f". Положите их в папку {FONTS_DIR}."
    )


def main() -> int:
    configure_stdio()
    print_banner()
    try:
        check_weasyprint()
    except RuntimeError as error:
        print(f"\n{error}")
        return 1

    data_files = list_files(DATA_DIR, (".csv", ".json"))
    templates = list_files(TEMPLATES_DIR, (".html", ".htm"))

    print_options("Доступные файлы с данными:", [path.name for path in data_files])
    print_options("Доступные шаблоны:", [path.name for path in templates])

    if not data_files:
        print(f"\nПоложите CSV или JSON в папку {DATA_DIR}.")
        return 1
    if not templates:
        print(f"\nПоложите HTML-шаблон в папку {TEMPLATES_DIR}.")
        return 1

    data_path = data_files[prompt_choice("Номер файла данных", len(data_files))]
    template_path = templates[prompt_choice("Номер шаблона", len(templates))]

    try:
        invoices = [invoice for invoice in load_invoices(data_path) if invoice.items]
    except (OSError, ValueError, json.JSONDecodeError, pd.errors.ParserError) as error:
        print(f"\nНе удалось прочитать {data_path.name}: {error}")
        return 1

    print_options(
        f"Доступные чеки в {data_path.name}:",
        [invoice_label(invoice) for invoice in invoices],
    )
    if not invoices:
        print("\nВ выбранном файле нет чеков с товарными строками.")
        return 1

    invoice = invoices[prompt_choice("Номер чека (invoice id)", len(invoices))]

    try:
        ensure_fonts()
        template_html = template_path.read_text(encoding="utf-8-sig")
        document = render_html(template_html, invoice)
        output_path = OUTPUT_DIR / f"nakladnaya-{safe_filename(invoice.invoice_id)}.pdf"
        write_pdf(document, template_path, output_path)
    except (OSError, ValueError, RuntimeError) as error:
        print(f"\nНе удалось создать PDF:\n{error}")
        return 1

    print(f"\nPDF сохранён: {output_path}")
    try:
        open_pdf(output_path)
        print("Документ открыт в системной программе.")
    except OSError as error:
        print(f"Не удалось открыть PDF автоматически: {error}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nОтменено.")
        raise SystemExit(130) from None
