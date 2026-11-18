import io
import os
import re
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from models import STATUS_CANCELLED, STATUS_FINAL, Contract, Declaration, db

HEADER_MARK = "شماره کوتاژ"
SUMMARY_HEADERS = {"ردیف", "شماره قرارداد"}


def cell_str(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value).rstrip("0").rstrip(".")
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if text.startswith("="):
        return ""
    return text


def cell_float(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.startswith("="):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_contract_key(number):
    text = cell_str(number)
    if text.endswith("/1"):
        return text[:-2]
    return text


def looks_like_contract_number(value):
    text = cell_str(value)
    if not text:
        return False
    return bool(re.fullmatch(r"140\d{6,9}(?:/\d+)?", text) or re.fullmatch(r"\d{9,12}(?:/\d+)?", text))


def looks_like_row_index(value):
    text = cell_str(value)
    if not text:
        return False
    return bool(re.fullmatch(r"\d{1,3}", text))


def detect_status(text):
    raw = cell_str(text)
    if "ابطال" in raw:
        return STATUS_CANCELLED
    if "نهایی" in raw:
        return STATUS_FINAL
    if raw:
        return raw
    return "در انتظار تکمیل"


def parse_slash_numbers(text):
    if not text:
        return []
    parts = re.split(r"[\\/|،,]+", str(text))
    return [p.strip() for p in parts if p.strip() and p.strip() != "-"]


def iter_sheet_rows(path):
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    max_col = min(ws.max_column or 9, 12)
    rows = []
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max_col), 1):
        values = [cell.value for cell in row]
        while len(values) < 9:
            values.append(None)
        rows.append((idx, values))
    return rows, ws.cell(1, 1).value


def parse_summary_contracts(rows):
    contracts = []
    header_row = None
    for idx, values in rows:
        a, b = cell_str(values[0]), cell_str(values[1])
        if a == "ردیف" and "قرارداد" in b:
            header_row = idx
            continue
        if header_row is None:
            continue
        if not looks_like_contract_number(b):
            if contracts and (not b and not cell_str(values[3])):
                break
            continue
        quantity = cell_float(values[4])
        if quantity is None:
            continue
        contracts.append(
            {
                "row": idx,
                "contract_number": cell_str(values[1]),
                "contract_date": cell_str(values[2]),
                "customer_name": cell_str(values[3]),
                "contract_quantity": quantity,
                "product_type": cell_str(values[5]),
                "delivery_location": cell_str(values[6]),
                "listed_declarations": parse_slash_numbers(values[7]),
                "excel_remaining": cell_float(values[8]),
            }
        )
    return contracts


def parse_declaration_blocks(rows):
    blocks = []
    current = None
    expecting_header = False
    unassigned = []

    def close_block():
        nonlocal current
        if current and current.get("declarations"):
            blocks.append(current)
        current = None

    for idx, values in rows:
        a, b, c, d, e = (
            cell_str(values[0]),
            cell_str(values[1]),
            values[2],
            cell_str(values[3]),
            cell_str(values[4]),
        )
        if idx < 26:
            continue
        if a == HEADER_MARK or (not a and cell_str(values[2]) == "تناژ"):
            if current is None:
                current = {
                    "contract_number": None,
                    "contract_date": "",
                    "customer_name": "",
                    "contract_quantity": None,
                    "product_type": "",
                    "delivery_location": "",
                    "declarations": [],
                    "source_row": idx,
                    "unassigned": True,
                }
            expecting_header = False
            continue
        if looks_like_contract_number(b) and (looks_like_row_index(a) or not a) and cell_float(values[4]) is not None:
            close_block()
            current = {
                "contract_number": cell_str(values[1]),
                "contract_date": cell_str(values[2]),
                "customer_name": cell_str(values[3]),
                "contract_quantity": cell_float(values[4]),
                "product_type": cell_str(values[5]),
                "delivery_location": cell_str(values[6]),
                "declarations": [],
                "source_row": idx,
                "unassigned": False,
            }
            expecting_header = True
            continue
        if current is None:
            continue
        decl_no = a
        tonnage = cell_float(c)
        if decl_no and not looks_like_contract_number(decl_no) and re.fullmatch(r"\d{4,8}", decl_no) and tonnage is not None:
            rec = {
                "declaration_number": decl_no,
                "tonnage": tonnage,
                "recipient": d,
                "description": e,
                "dispatch_reference_number": cell_str(values[7]),
                "status": detect_status(e),
                "source_row": idx,
            }
            current["declarations"].append(rec)
            continue
        if isinstance(c, str) and str(c).startswith("="):
            close_block()
            continue
        if not a and not cell_str(b) and not cell_str(d) and tonnage is None:
            if current and current["declarations"]:
                close_block()

    close_block()

    assigned = []
    for block in blocks:
        if block["contract_number"]:
            assigned.append(block)
        else:
            unassigned.extend(block["declarations"])
    return assigned, unassigned


def match_unassigned(unassigned, summary_contracts, assigned_numbers):
    index = defaultdict(list)
    for item in summary_contracts:
        for num in item["listed_declarations"]:
            index[num].append(item["contract_number"])
    matched = defaultdict(list)
    flagged = []
    for rec in unassigned:
        num = rec["declaration_number"]
        if num in assigned_numbers:
            continue
        owners = list(dict.fromkeys(index.get(num, [])))
        if len(owners) == 1:
            matched[owners[0]].append(rec)
        else:
            rec = dict(rec)
            rec["needs_review"] = True
            rec["review_note"] = "ارتباط این کوتاژ با قرارداد در فایل اکسل مبهم است و نیاز به بررسی دستی دارد."
            flagged.append(rec)
    return matched, flagged


def import_workbook(path, company_id):
    rows, title = iter_sheet_rows(path)
    summary = parse_summary_contracts(rows)
    blocks, unassigned = parse_declaration_blocks(rows)

    by_key = {}
    for item in summary:
        key = normalize_contract_key(item["contract_number"])
        by_key[key] = item

    created_contracts = 0
    created_declarations = 0
    flagged_count = 0
    contract_map = {}

    def get_or_create_contract(info, extra_review=""):
        nonlocal created_contracts, flagged_count
        number = cell_str(info["contract_number"])
        existing = Contract.query.filter_by(company_id=company_id, contract_number=number, is_deleted=False).first()
        if existing:
            contract_map[normalize_contract_key(number)] = existing
            return existing
        key = normalize_contract_key(number)
        sibling = contract_map.get(key)
        if sibling:
            return sibling
        review = extra_review or info.get("review_note", "")
        contract = Contract(
            company_id=company_id,
            contract_number=number,
            contract_date=info.get("contract_date") or "",
            customer_name=(info.get("customer_name") or "").strip(),
            contract_quantity=info.get("contract_quantity") or 0,
            product_type=(info.get("product_type") or "").strip(),
            delivery_location=(info.get("delivery_location") or "").strip(),
            needs_review=bool(review),
            review_note=review,
        )
        db.session.add(contract)
        db.session.flush()
        created_contracts += 1
        if review:
            flagged_count += 1
        contract_map[key] = contract
        return contract

    for item in summary:
        get_or_create_contract(item)

    assigned_decl_numbers = set()
    for block in blocks:
        key = normalize_contract_key(block["contract_number"])
        info = by_key.get(key)
        if info is None:
            review = "این قرارداد فقط در بخش جزئیات کوتاژ دیده شد و در جدول اصلی اکسل وجود نداشت."
            contract = get_or_create_contract(block, extra_review=review)
        else:
            preferred = dict(info)
            preferred["contract_number"] = info["contract_number"]
            contract = get_or_create_contract(preferred)
        for rec in block["declarations"]:
            if Declaration.query.filter_by(
                contract_id=contract.id, declaration_number=rec["declaration_number"], is_deleted=False
            ).first():
                continue
            decl = Declaration(
                contract_id=contract.id,
                declaration_number=rec["declaration_number"],
                tonnage=rec["tonnage"] or 0,
                recipient=rec.get("recipient") or "",
                description=rec.get("description") or "",
                dispatch_reference_number=rec.get("dispatch_reference_number") or "",
                status=rec.get("status") or STATUS_FINAL,
            )
            db.session.add(decl)
            created_declarations += 1
            assigned_decl_numbers.add(rec["declaration_number"])

    matched, flagged = match_unassigned(unassigned, summary, assigned_decl_numbers)
    for contract_number, recs in matched.items():
        contract = get_or_create_contract({"contract_number": contract_number, "contract_quantity": 0})
        for rec in recs:
            if Declaration.query.filter_by(
                contract_id=contract.id, declaration_number=rec["declaration_number"], is_deleted=False
            ).first():
                continue
            db.session.add(
                Declaration(
                    contract_id=contract.id,
                    declaration_number=rec["declaration_number"],
                    tonnage=rec["tonnage"] or 0,
                    recipient=rec.get("recipient") or "",
                    description=rec.get("description") or "",
                    dispatch_reference_number=rec.get("dispatch_reference_number") or "",
                    status=rec.get("status") or STATUS_FINAL,
                )
            )
            created_declarations += 1

    db.session.flush()
    for item in summary:
        contract = contract_map.get(normalize_contract_key(item["contract_number"]))
        if not contract:
            continue
        assigned = {d.declaration_number for d in contract.active_declarations()}
        listed = set(item.get("listed_declarations") or [])
        notes = []
        missing = listed - assigned
        extra = assigned - listed
        if missing:
            notes.append("شماره کوتاژهای جدول اصلی که جزئیات تناژ ندارند: " + "/".join(sorted(missing)))
        if extra:
            notes.append("کوتاژهایی که در جزئیات هستند ولی در فهرست اصلی نیستند: " + "/".join(sorted(extra)))
        excel_rem = item.get("excel_remaining")
        if excel_rem is not None and abs((contract.remaining_tonnage or 0) - excel_rem) > 0.01:
            notes.append(
                f"مانده تناژ فایل اکسل ({excel_rem}) با محاسبه سامانه ({contract.remaining_tonnage}) متفاوت است."
            )
        if notes:
            contract.needs_review = True
            contract.review_note = (contract.review_note + " " + " ".join(notes)).strip()
            flagged_count += 1

    if flagged:
        review_contract = Contract.query.filter_by(
            company_id=company_id, contract_number="REVIEW-UNASSIGNED", is_deleted=False
        ).first()
        if review_contract is None:
            review_contract = Contract(
                company_id=company_id,
                contract_number="REVIEW-UNASSIGNED",
                contract_date="",
                customer_name="نیاز به بررسی دستی",
                contract_quantity=0,
                product_type="",
                delivery_location="",
                needs_review=True,
                review_note="کوتاژهایی که ارتباط آن‌ها با قرارداد در فایل اکسل مبهم بود.",
            )
            db.session.add(review_contract)
            db.session.flush()
            created_contracts += 1
            flagged_count += 1
        for rec in flagged:
            db.session.add(
                Declaration(
                    contract_id=review_contract.id,
                    declaration_number=rec["declaration_number"],
                    tonnage=rec["tonnage"] or 0,
                    recipient=rec.get("recipient") or "",
                    description=rec.get("description") or "",
                    dispatch_reference_number=rec.get("dispatch_reference_number") or "",
                    status=rec.get("status") or STATUS_FINAL,
                    needs_review=True,
                    review_note=rec.get("review_note") or "",
                )
            )
            created_declarations += 1
            flagged_count += 1

    db.session.commit()
    return {
        "title": cell_str(title),
        "contracts": created_contracts,
        "declarations": created_declarations,
        "flagged": flagged_count,
    }


def style_header(cell, fill_hex="1F4E46"):
    cell.font = Font(name="Tahoma", bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", fgColor=fill_hex)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color="C4B59A"),
        right=Side(style="thin", color="C4B59A"),
        top=Side(style="thin", color="C4B59A"),
        bottom=Side(style="thin", color="C4B59A"),
    )


def style_cell(cell, bold=False):
    cell.font = Font(name="Tahoma", bold=bold, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color="D9D0C1"),
        right=Side(style="thin", color="D9D0C1"),
        top=Side(style="thin", color="D9D0C1"),
        bottom=Side(style="thin", color="D9D0C1"),
    )


def write_text(cell, value, bold=False):
    cell.value = "" if value is None else str(value)
    cell.number_format = "@"
    style_cell(cell, bold=bold)


def export_contracts_workbook(company, contracts, title=None):
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    ws.title = "گزارش کوتاژ"
    title = title or f"گزارش کوتاژ و اظهار {company.name}"
    ws.merge_cells("A1:I1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Tahoma", bold=True, size=14, color="1F4E46")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = [
        "ردیف",
        "شماره قرارداد",
        "تاریخ قرارداد",
        "نام کامل مشتری",
        "مقدار",
        "نام محصول",
        "محل تحویل",
        "شماره کوتاژ ها",
        "مانده تناژ",
    ]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(2, col, header)
        style_header(cell)
    ws.row_dimensions[2].height = 22

    for i, contract in enumerate(contracts, 1):
        ws.cell(2 + i, 1, i)
        style_cell(ws.cell(2 + i, 1))
        write_text(ws.cell(2 + i, 2), contract.contract_number)
        write_text(ws.cell(2 + i, 3), contract.contract_date)
        write_text(ws.cell(2 + i, 4), contract.customer_name)
        ws.cell(2 + i, 5, contract.contract_quantity)
        style_cell(ws.cell(2 + i, 5))
        write_text(ws.cell(2 + i, 6), contract.product_type)
        write_text(ws.cell(2 + i, 7), contract.delivery_location)
        write_text(ws.cell(2 + i, 8), contract.declaration_numbers_text)
        rem = ws.cell(2 + i, 9, contract.remaining_tonnage)
        style_cell(rem, bold=True)
        rem.font = Font(name="Tahoma", bold=True, color="8A1C1C")

    row = 4 + len(contracts)
    ws.cell(row, 4, "مانده تناژها و کوتاژها")
    ws.cell(row, 4).font = Font(name="Tahoma", bold=True, size=12, color="1F4E46")
    groups = defaultdict(float)
    for contract in contracts:
        key = (contract.product_type or "", contract.delivery_location or "")
        groups[key] += contract.remaining_tonnage or 0
    row += 1
    ws.cell(row, 4, "نام محصول")
    ws.cell(row, 5, "محل تحویل")
    ws.cell(row, 6, "مانده تناژ")
    for col in range(4, 7):
        style_header(ws.cell(row, col), "3D6B62")
    start = row + 1
    for (product, location), total in sorted(groups.items(), key=lambda x: x[0][0]):
        row += 1
        ws.cell(row, 4, product)
        ws.cell(row, 5, location)
        ws.cell(row, 6, total)
        for col in range(4, 7):
            style_cell(ws.cell(row, col))
    row += 1
    ws.cell(row, 4, "جمع کل")
    ws.cell(row, 6, sum(groups.values()))
    style_header(ws.cell(row, 4), "C4A35A")
    style_header(ws.cell(row, 5), "C4A35A")
    style_header(ws.cell(row, 6), "C4A35A")

    row += 3
    for i, contract in enumerate(contracts, 1):
        decls = contract.active_declarations()
        ws.cell(row, 1, i)
        write_text(ws.cell(row, 2), contract.contract_number, bold=True)
        write_text(ws.cell(row, 3), contract.contract_date, bold=True)
        write_text(ws.cell(row, 4), contract.customer_name, bold=True)
        ws.cell(row, 5, contract.contract_quantity)
        write_text(ws.cell(row, 6), contract.product_type, bold=True)
        ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=8)
        write_text(ws.cell(row, 7), contract.delivery_location, bold=True)
        for col in range(1, 9):
            style_cell(ws.cell(row, col), bold=True)
            ws.cell(row, col).fill = PatternFill("solid", fgColor="EFE6D6")
        row += 1
        for col, header in enumerate(["شماره کوتاژ", "", "تناژ", "برای شخص", "توضیحات", "", "", "شماره حواله"], 1):
            cell = ws.cell(row, col, header)
            if header:
                style_header(cell, "5C4A32")
        row += 1
        sum_start = row
        for decl in decls:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
            write_text(ws.cell(row, 1), decl.declaration_number)
            ws.cell(row, 3, decl.tonnage)
            style_cell(ws.cell(row, 3))
            write_text(ws.cell(row, 4), decl.recipient)
            write_text(ws.cell(row, 5), decl.status if not decl.description else decl.description)
            write_text(ws.cell(row, 8), decl.dispatch_reference_number)
            for col in range(1, 9):
                style_cell(ws.cell(row, col))
            if decl.status == STATUS_CANCELLED:
                ws.cell(row, 5).font = Font(name="Tahoma", color="8A1C1C", bold=True)
            row += 1
        if decls:
            finalized_total = sum(d.tonnage or 0 for d in decls if d.status == STATUS_FINAL)
            ws.cell(row, 3, finalized_total)
            style_cell(ws.cell(row, 3), bold=True)
            row += 2
        else:
            row += 1

    widths = [12, 18, 16, 28, 14, 38, 22, 42, 14]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def default_excel_paths():
    base = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
    return {
        "شرکت گچ خراسان": os.path.join(base, "gach_khorasan.xlsx"),
        "شرکت تولیدی معدنی آتشکده خراسان": os.path.join(base, "atashkadeh.xlsx"),
    }
