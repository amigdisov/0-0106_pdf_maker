#!/usr/bin/env python3
"""Генератор PDF-накладных из CSV/JSON и HTML-шаблонов.

Как программа работает, если смотреть сверху:

1. Ищет файлы с данными в папке data (CSV и JSON) и HTML-шаблоны в templates.
2. Показывает нумерованное меню и просит выбрать файл, шаблон и чек (invoice id).
3. Читает выбранный файл и собирает из него список накладных.
4. Подставляет товары выбранной накладной в HTML (вместо {{ product }} и других меток).
5. Библиотека WeasyPrint превращает этот HTML в PDF и кладёт файл в папку output.
6. PDF открывается в программе, которая на компьютере назначена для PDF.

CSV читается через pandas, JSON — через стандартный модуль json.
Кириллица в PDF рисуется шрифтом DejaVu Sans из папки fonts.
"""

# Позволяет писать современные подсказки типов (list[Item], str | None)
# даже на более старых версиях Python. На результат работы не влияет.
from __future__ import annotations

import html  # html.escape защищает текст товара, если в нём есть символы <, >, &
import json  # стандартная библиотека: читает и разбирает JSON
import math  # нужна, чтобы отличить пустое число NaN от обычного float
import os  # переменные окружения и открытие файла в Windows
import platform  # узнать, Windows это, macOS или Linux
import re  # регулярные выражения: поиск меток {{ ... }} в шаблоне
import subprocess  # запуск системной команды open / xdg-open
import sys  # доступ к консоли (stdin/stdout) и коду возврата
from dataclasses import dataclass, field  # короткая запись классов для данных
from datetime import datetime  # разбор и вывод даты накладной
from pathlib import Path  # удобные пути к файлам, без ручной склейки строк

import pandas as pd  # таблица CSV: чтение, имена колонок, строки как словари

# Папка, в которой лежит этот скрипт. Не зависит от того, откуда его запустили.
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"  # исходные CSV и JSON
TEMPLATES_DIR = BASE_DIR / "templates"  # HTML-шаблоны накладной
OUTPUT_DIR = BASE_DIR / "output"  # сюда сохраняются готовые PDF
FONTS_DIR = BASE_DIR / "fonts"  # DejaVu Sans: обычный и жирный

# Ставка НДС 20%. В накладной сумма уже включает налог,
# поэтому налог считается «изнутри» суммы, а не сверху.
VAT_RATE = 0.20

# Ищет в HTML метки вида {{ product }} или {{ invoice_id }}.
# \s* разрешает пробелы внутри скобок: {{  total  }} тоже подойдёт.
# Скобки вокруг имени сохраняют его: потом это match.group(1).
PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")

# Ищет одну строку таблицы, в которой есть метка {{ product }}.
# Шаблон хранит её один раз, а скрипт копирует её для каждого товара.
# re.DOTALL нужен, чтобы точка совпала и с переносом строки.
# re.IGNORECASE — чтобы <TR> и <tr> считались одним и тем же.
ITEM_ROW_RE = re.compile(
    r"<tr\b[^>]*>.*?\{\{\s*product\s*\}\}.*?</tr>",
    re.IGNORECASE | re.DOTALL,
)

# Одни и те же данные в файлах могут называться по-разному.
# Скрипт перебирает эти имена и берёт первое непустое.
ID_KEYS = ("invoice_id", "invoiceid", "invoice", "id", "номер")
DATE_KEYS = ("date", "invoice_date", "дата")
PRODUCT_KEYS = ("product", "name", "товар", "наименование")
PRICE_KEYS = ("price", "цена")
QUANTITY_KEYS = ("quantity", "qty", "количество", "кол-во", "кол_во")


@dataclass
class Item:
    """Одна товарная строка накладной.

    dataclass сам создаёт __init__: можно писать Item("Яблоки", 120.5, 10)
    и не описывать конструктор вручную.
    """

    product: str
    price: float
    quantity: float

    @property
    def line_sum(self) -> float:
        """Сумма строки: цена умножить на количество.

        @property позволяет писать item.line_sum как поле, хотя это расчёт.
        """
        return self.price * self.quantity


@dataclass
class Invoice:
    """Одна накладная: номер, дата и список товаров.

    default_factory=list нужен, чтобы у каждой накладной был свой список.
    Если написать items: list = [], все объекты делили бы один и тот же список.
    """

    invoice_id: str
    items: list[Item] = field(default_factory=list)
    date: str = ""

    @property
    def total(self) -> float:
        """Итог по всем строкам."""
        return sum(item.line_sum for item in self.items)

    @property
    def vat(self) -> float:
        """НДС 20%, уже включённый в итог.

        Формула «налог внутри суммы»: сумма * 0.20 / 1.20.
        Пример: 120 рублей с НДС 20% содержат 20 рублей налога, а не 24.
        """
        return self.total * VAT_RATE / (1 + VAT_RATE)


def configure_stdio() -> None:
    """Включает UTF-8 в консоли Windows.

    Без этого русские буквы в меню могут превратиться в кракозябры.
    На macOS и Linux консоль обычно уже в UTF-8, поэтому функция сразу выходит.
    """
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        # У старых объектов потока метода reconfigure может не быть.
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8")
        except Exception:
            # Если кодировку сменить нельзя, продолжаем с тем, что есть.
            pass


def prepare_weasyprint_env() -> None:
    """На Windows подсказывает WeasyPrint, где лежат DLL библиотеки Pango.

    WeasyPrint сам по себе — Python-пакет, но рисует текст через системные
    библиотеки Pango. На Windows их обычно ставят через MSYS2. Переменная
    WEASYPRINT_DLL_DIRECTORIES — это список папок, в которых нужно искать DLL.

    Если переменная уже задана в системе, скрипт её не трогает.
    Если папка C:\\msys64\\ucrt64\\bin существует, путь подставится сам.
    """
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
        # os.pathsep на Windows — точка с запятой, на macOS/Linux — двоеточие.
        os.environ["WEASYPRINT_DLL_DIRECTORIES"] = os.pathsep.join(existing)


def as_text(value: object) -> str:
    """Превращает значение из таблицы в обычную строку без мусора.

    pandas часто отдаёт числа как float. Тогда номер 7 приходит как 7.0,
    а пустая ячейка — как NaN («не число»). Здесь 7.0 становится "7",
    а NaN — пустой строкой.
    """
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
    """Достаёт число из ячейки. Пустое или нечисловое значение даёт None.

    Принимает и «1 205,50», и «1205.50»: пробелы убираются,
    запятая заменяется на точку, потому что float() понимает только точку.
    """
    text = as_text(value).replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def format_date(value: object) -> str:
    """Приводит дату к виду ДД.ММ.ГГГГ.

    Понимает 2026-10-07, 07.10.2026 и 07/10/2026.
    Если формат незнакомый, возвращает текст как есть, а не падает с ошибкой.
    """
    text = as_text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            # Берём первые 10 символов, чтобы отсечь время, если оно пришло вместе с датой.
            return datetime.strptime(text[:10], fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def format_money(value: float) -> str:
    """Форматирует деньги: 6682.0 -> «6 682.00».

    Сначала f-строка ставит запятую как разделитель тысяч (так умеет Python),
    потом запятая меняется на пробел — так принято в русской вёрстке.
    """
    return f"{value:,.2f}".replace(",", " ")


def format_quantity(value: float) -> str:
    """Количество: целое показывает без дроби (10), дробное — как деньги (1.50)."""
    if float(value).is_integer():
        return str(int(value))
    return format_money(value)


def normalize_key(key: object) -> str:
    """Приводит имя поля к одному виду: «Invoice ID» и «invoice-id» -> invoice_id.

    Так CSV с разными заголовками читается одними и теми же списками ключей.
    """
    return str(key).strip().lower().replace(" ", "_").replace("-", "_")


def first_present(source: dict, keys: tuple[str, ...]) -> object:
    """Возвращает первое непустое поле из словаря по списку возможных имён.

    Например, для номера чека перебираются invoice_id, id, номер и так далее.
    Если ничего нет, возвращает None.
    """
    normalized = {normalize_key(key): value for key, value in source.items()}
    for key in keys:
        if key in normalized and as_text(normalized[key]):
            return normalized[key]
    return None


def item_from_mapping(source: dict) -> Item | None:
    """Собирает одну товарную строку из словаря CSV/JSON.

    Строка без названия, цены или количества пропускается (вернётся None).
    Так битая строка не ломает всю накладную.
    """
    product = as_text(first_present(source, PRODUCT_KEYS))
    price = parse_number(first_present(source, PRICE_KEYS))
    quantity = parse_number(first_present(source, QUANTITY_KEYS))
    if not product or price is None or quantity is None:
        return None
    return Item(product=product, price=price, quantity=quantity)


def invoice_from_mapping(source: dict, fallback_id: str) -> Invoice:
    """Собирает накладную из JSON-объекта, у которого товары лежат в списке items.

    fallback_id используется, если в объекте нет номера: тогда номер берётся
    из имени файла, чтобы в меню всё равно было что выбрать.
    """
    invoice_id = as_text(first_present(source, ID_KEYS)) or fallback_id
    date = format_date(first_present(source, DATE_KEYS))
    # В разных файлах список товаров может называться items, products или lines.
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
    """Группирует плоские строки в накладные по invoice id.

    Так устроен CSV: одна накладная занимает несколько строк, по строке на товар.
    Строки с одним и тем же номером склеиваются в один объект Invoice.
    Если колонки с номером нет, все строки становятся одной накладной
    с номером, равным имени файла (fallback_id).

    Отдельный список order хранит порядок номеров, как они встретились в файле.
    Обычный dict в новых Python тоже помнит порядок, но явный список проще читать.
    """
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
        # Дата могла быть пустой в первой строке и заполненной в следующей.
        if not invoice.date:
            invoice.date = format_date(first_present(row, DATE_KEYS))
        item = item_from_mapping(row)
        if item is not None:
            invoice.items.append(item)
    return [grouped[invoice_id] for invoice_id in order]


def load_csv(path: Path) -> list[Invoice]:
    """Читает CSV через pandas и возвращает список накладных.

    encoding utf-8-sig понимает файлы, которые Excel сохранил с невидимым
    символом BOM в начале. Без этого первая колонка могла бы называться
    «\\ufeffinvoice_id» вместо «invoice_id».
    """
    frame = pd.read_csv(path, encoding="utf-8-sig")
    frame.columns = [normalize_key(column) for column in frame.columns]
    # orient="records" даёт список словарей: одна строка таблицы — один словарь.
    rows = frame.to_dict(orient="records")
    # path.stem — имя файла без расширения. Для csvsource.csv это «csvsource».
    return invoices_from_rows(rows, fallback_id=path.stem)


def load_json(path: Path) -> list[Invoice]:
    """Читает JSON стандартной библиотекой и приводит его к списку накладных.

    Поддерживаются три формы, потому что файлы пишут по-разному:

    1. Список накладных, у каждой есть поле items.
    2. Объект {"invoices": [ ... ]} или объект «номер -> накладная».
    3. Плоский список строк, как в CSV: номер и товар в одной записи.
    """
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    fallback_id = path.stem

    if isinstance(payload, dict):
        invoices = payload.get("invoices")
        if isinstance(invoices, list):
            # Форма {"invoices": [ {...}, {...} ]}.
            payload = invoices
        elif payload and all(isinstance(value, dict) for value in payload.values()):
            # Форма {"2026-0007": {"date": "...", "items": [...]}}.
            # Ключ словаря становится номером, если внутри номера нет.
            documents = []
            for key, value in payload.items():
                document = dict(value)
                document.setdefault("invoice_id", key)
                documents.append(document)
            payload = documents
        else:
            # Один объект накладной без обёртки.
            payload = [payload]

    if not isinstance(payload, list):
        raise ValueError(f"Неподдерживаемая структура JSON: {path.name}")

    if not payload:
        return []

    # Плоский список строк не содержит вложенного items/products/lines,
    # зато в каждой записи есть название товара.
    looks_like_flat_rows = all(
        isinstance(item, dict) and not any(key in item for key in ("items", "products", "lines"))
        for item in payload
    )
    if looks_like_flat_rows and any(first_present(item, PRODUCT_KEYS) is not None for item in payload):
        return invoices_from_rows(payload, fallback_id=fallback_id)

    invoices: list[Invoice] = []
    for index, item in enumerate(payload, start=1):
        if not isinstance(item, dict):
            continue
        invoices.append(invoice_from_mapping(item, fallback_id=f"{fallback_id}-{index}"))
    return invoices


def load_invoices(path: Path) -> list[Invoice]:
    """Выбирает способ чтения по расширению файла."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return load_csv(path)
    if suffix == ".json":
        return load_json(path)
    raise ValueError(f"Неподдерживаемый формат: {path.name}")


def list_files(directory: Path, suffixes: tuple[str, ...]) -> list[Path]:
    """Возвращает файлы папки с нужными расширениями, по алфавиту.

    Скрытые файлы (имя начинается с точки) пропускаются.
    Словарь found убирает дубли, если система отдала один файл дважды.
    """
    if not directory.is_dir():
        return []
    found: dict[Path, Path] = {}
    for path in directory.iterdir():
        if path.is_file() and path.suffix.lower() in suffixes and not path.name.startswith("."):
            found[path.resolve()] = path
    return sorted(found.values(), key=lambda item: item.name.lower())


def fill_placeholders(template: str, values: dict[str, str]) -> str:
    """Заменяет {{ имя }} на значение из словаря.

    Если такого имени в словаре нет, метка остаётся как была.
    Это удобно: в строке товара заменяются product и price,
    а {{ total }} ждёт следующего прохода, когда итог уже посчитан.
    """

    def replace(match: re.Match[str]) -> str:
        return values.get(match.group(1), match.group(0))

    return PLACEHOLDER_RE.sub(replace, template)


def apply_fonts(template_html: str) -> str:
    """Подключает DejaVu Sans так, чтобы кириллица попала в PDF.

    В шаблоне шрифт указан относительным путём ../fonts/DejaVuSans.ttf.
    Здесь путь заменяется на полный file://..., потому что WeasyPrint
    надёжнее открывает абсолютный адрес, особенно если в папке проекта
    есть пробелы или русские буквы.

    Если в шаблоне вообще нет @font-face, правило дописывается само.
    Сначала заменяется жирный файл: его имя длиннее и содержит имя обычного.
    """
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
    """Собирает готовый HTML одной накладной.

    Строка таблицы с {{ product }} в шаблоне записана один раз.
    Функция вырезает её, заполняет копиями по числу товаров и вставляет обратно.
    Потом отдельно подставляет номер, дату, итог, НДС и число позиций.
    html.escape нужен, чтобы название вроде «Сыр <20%>» не сломало разметку.
    """
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

    # Склеиваем: HTML до строки + все товарные строки + HTML после строки.
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
    """Делает из номера чека безопасное имя файла.

    Пробелы и символы вроде / \\ : * заменяются на подчёркивание,
    чтобы Windows и macOS приняли имя. Буквы, цифры, точка и дефис остаются.
    """
    cleaned = re.sub(r"[^\w.\-]+", "_", invoice_id, flags=re.UNICODE).strip("._")
    return cleaned or "invoice"


def write_pdf(html_document: str, template_path: Path, output_path: Path) -> None:
    """Превращает HTML в PDF через WeasyPrint и сохраняет файл.

    Импорт стоит внутри функции, а не в начале файла: если библиотек Pango нет,
    программа успевает показать понятную подсказку, а не падает на первой строке.

    base_url — папка шаблона. Относительные картинки и стили ищутся от неё.
    """
    prepare_weasyprint_env()
    try:
        from weasyprint import HTML
    except OSError as error:
        raise RuntimeError(weasyprint_help(error)) from error

    output_path.parent.mkdir(parents=True, exist_ok=True)
    HTML(string=html_document, base_url=str(template_path.parent)).write_pdf(output_path)


def weasyprint_help(error: BaseException) -> str:
    """Текст подсказки, если WeasyPrint не смог найти системные библиотеки."""
    lines = [
        "Не удалось загрузить WeasyPrint. Для PDF нужны библиотеки Pango.",
        f"Подробности: {error}",
    ]
    if platform.system() == "Windows":
        lines.extend(
            [
                "Windows: установите MSYS2 (https://www.msys2.org/) и в оболочке UCRT64 выполните:",
                "  pacman -S mingw-w64-ucrt-x86_64-pango",
                "Если Pango лежит в C:\\msys64\\ucrt64\\bin, скрипт найдёт его сам.",
                "Иначе перед запуском задайте переменную в том же окне, где запускаете Python.",
                "PowerShell:",
                '  $env:WEASYPRINT_DLL_DIRECTORIES = "C:\\msys64\\ucrt64\\bin"',
                "cmd.exe:",
                r"  set WEASYPRINT_DLL_DIRECTORIES=C:\msys64\ucrt64\bin",
            ]
        )
    elif platform.system() == "Darwin":
        lines.append("macOS: brew install pango")
    return "\n".join(lines)


def open_pdf(path: Path) -> None:
    """Открывает готовый PDF в программе по умолчанию.

    У каждой системы своя команда:
    Windows — os.startfile, macOS — open, Linux — xdg-open.
    """
    system = platform.system()
    if system == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
        return
    if system == "Darwin":
        subprocess.run(["open", str(path)], check=False)
        return
    subprocess.run(["xdg-open", str(path)], check=False)


def print_banner() -> None:
    """Печатает заголовок меню."""
    print()
    print("=" * 48)
    print("  Генератор PDF-накладных")
    print("=" * 48)


def print_options(title: str, labels: list[str]) -> None:
    """Печатает нумерованный список. Нумерация с 1, как привычно человеку."""
    print()
    print(title)
    if not labels:
        print("  (ничего не найдено)")
        return
    for index, label in enumerate(labels, start=1):
        print(f"  {index}. {label}")


def prompt_choice(label: str, count: int) -> int:
    """Спрашивает номер пункта, пока пользователь не введёт правильный.

    Возвращает индекс с нуля: пункт «1» в меню — это элемент 0 в списке Python.
    Цикл не заканчивается, пока ввод не будет числом из диапазона.
    """
    while True:
        raw = input(f"\n{label}: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= count:
            return int(raw) - 1
        print(f"Введите номер от 1 до {count}.")


def invoice_label(invoice: Invoice) -> str:
    """Строка меню для одного чека: номер, сколько позиций и дата."""
    details = [f"{len(invoice.items)} поз."]
    if invoice.date:
        details.append(invoice.date)
    return f"{invoice.invoice_id}  ({', '.join(details)})"


def check_weasyprint() -> None:
    """Проверяет WeasyPrint до меню, чтобы не выбирать чек впустую.

    Импорт HTML заставляет библиотеку загрузить Pango. Если DLL нет,
    возникает OSError, и мы превращаем его в понятный RuntimeError.
    """
    prepare_weasyprint_env()
    try:
        from weasyprint import HTML  # noqa: F401
    except OSError as error:
        raise RuntimeError(weasyprint_help(error)) from error


def ensure_fonts() -> None:
    """Проверяет, что оба файла DejaVu Sans лежат в папке fonts.

    Нужны обычное и жирное начертание: заголовки в шаблоне набраны жирным.
    Без жирного шрифта кириллица в заголовках может подмениться другим шрифтом.
    """
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
    """Точка входа: меню, чтение данных, сборка PDF.

    Возвращает 0, если PDF создан, и 1, если на каком-то шаге произошла ошибка.
    Этот код потом передаётся операционной системе через SystemExit.
    """
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
        # Накладные без товаров в меню не показываем: из них нечего печатать.
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
        # Файл уже сохранён. Не смогли только открыть его автоматически.
        print(f"Не удалось открыть PDF автоматически: {error}")
    return 0


if __name__ == "__main__":
    # Этот блок выполняется только при запуске «python main.py»,
    # и не выполняется, если файл импортируют из другого скрипта.
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        # Пользователь нажал Ctrl+C. 130 — обычный код «прервано с клавиатуры».
        print("\nОтменено.")
        raise SystemExit(130) from None
