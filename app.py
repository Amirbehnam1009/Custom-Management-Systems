import os
import secrets
import shutil
from datetime import datetime
from functools import wraps

from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from sqlalchemy import event
from sqlalchemy.engine import Engine
from werkzeug.utils import secure_filename

from models import (
    STATUS_CANCELLED,
    STATUS_FINAL,
    STATUS_PENDING,
    STATUSES,
    AuditLog,
    Company,
    Contract,
    Declaration,
    User,
    db,
)
from utils.excel_io import default_excel_paths, export_contracts_workbook, import_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, "instance")
BACKUP_DIR = os.path.join(INSTANCE_DIR, "backups")
UPLOAD_DIR = os.path.join(INSTANCE_DIR, "uploads")
os.makedirs(INSTANCE_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("APP_SECRET_KEY", "kotaj-local-dev-key")
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(INSTANCE_DIR, "kotaj.db").replace("\\", "/")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SEED_EXCEL"] = True
app.config["CSRF_DISABLE"] = False

db.init_app(app)


@event.listens_for(Engine, "connect")
def _enable_sqlite_foreign_keys(dbapi_connection, connection_record):
    try:
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.close()
    except Exception:
        pass

COMPANIES = [
    {"name": "شرکت تولیدی معدنی آتشکده خراسان", "short_name": "آتشکده"},
    {"name": "شرکت گچ خراسان", "short_name": "گچ خراسان"},
]


def current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    return db.session.get(User, user_id)


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not current_user():
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "error": "لطفاً وارد شوید."}), 401
            return redirect(url_for("login"))
        return fn(*args, **kwargs)

    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user:
            if request.path.startswith("/api/") and request.method not in {"GET", "HEAD"}:
                return jsonify({"ok": False, "error": "لطفاً وارد شوید."}), 401
            return redirect(url_for("login"))
        if not user.is_admin:
            if request.path.startswith("/api/") and request.content_type and "json" in (request.content_type or ""):
                return jsonify({"ok": False, "error": "فقط مدیر سیستم به این بخش دسترسی دارد."}), 403
            flash("فقط مدیر سیستم به این بخش دسترسی دارد.", "error")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)

    return wrapper


def csrf_token():
    token = session.get("_csrf")
    if not token:
        token = secrets.token_hex(16)
        session["_csrf"] = token
    return token


def _submitted_csrf_token():
    token = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token")
    if token:
        return token
    if request.is_json:
        payload = request.get_json(silent=True) or {}
        if isinstance(payload, dict):
            return payload.get("csrf_token")
    return None


def active_contract_or_404(contract_id):
    contract = db.session.get(Contract, contract_id)
    if not contract or contract.is_deleted:
        abort(404)
    return contract


def active_declaration_or_404(declaration_id):
    declaration = db.session.get(Declaration, declaration_id)
    if not declaration or declaration.is_deleted:
        abort(404)
    if declaration.contract and declaration.contract.is_deleted:
        abort(404)
    return declaration


def safe_excel_filename(filename):
    original = filename or "upload.xlsx"
    name = secure_filename(original)
    ext = os.path.splitext(original)[1].lower()
    if ext not in {".xlsx", ".xls"}:
        raise ValueError("فقط فایل اکسل با پسوند xlsx مجاز است.")
    if not name or name.lower() in {".xlsx", ".xls"}:
        name = datetime.now().strftime("import-%Y%m%d-%H%M%S") + ext
    if os.path.splitext(name)[1].lower() not in {".xlsx", ".xls"}:
        name += ext
    return name


def log_action(action, record_type="", record_id=None, details=""):
    user = current_user()
    db.session.add(
        AuditLog(
            user_id=user.id if user else None,
            action=action,
            record_type=record_type,
            record_id=record_id,
            details=details,
        )
    )
    db.session.commit()


def json_error(message, code=400):
    return jsonify({"ok": False, "error": message}), code


PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def parse_float(value, field_name):
    if value is None or str(value).strip() == "":
        raise ValueError(f"{field_name} الزامی است.")
    text = str(value).translate(PERSIAN_DIGITS).replace(",", "").replace("،", "").strip()
    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(f"{field_name} باید عدد باشد.") from exc
    if number < 0:
        raise ValueError(f"{field_name} نمی‌تواند منفی باشد.")
    return number


def apply_contract_filters(query):
    company_id = request.args.get("company_id", type=int)
    q = (request.args.get("q") or "").strip()
    product = (request.args.get("product") or "").strip()
    location = (request.args.get("location") or "").strip()
    customer = (request.args.get("customer") or "").strip()
    date_from = (request.args.get("date_from") or "").strip()
    date_to = (request.args.get("date_to") or "").strip()
    remaining = (request.args.get("remaining") or "").strip()
    status = (request.args.get("status") or "").strip()

    query = query.filter(Contract.is_deleted.is_(False))
    if company_id:
        query = query.filter(Contract.company_id == company_id)
    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Contract.contract_number.ilike(like),
                Contract.customer_name.ilike(like),
            )
        )
    if customer:
        query = query.filter(Contract.customer_name.ilike(f"%{customer}%"))
    if product:
        query = query.filter(Contract.product_type.ilike(f"%{product}%"))
    if location:
        query = query.filter(Contract.delivery_location.ilike(f"%{location}%"))
    if date_from:
        query = query.filter(Contract.contract_date >= date_from)
    if date_to:
        query = query.filter(Contract.contract_date <= date_to)
    contracts = query.order_by(Contract.id.desc()).all()
    if remaining == "has":
        contracts = [c for c in contracts if c.remaining_tonnage > 0]
    elif remaining == "zero":
        contracts = [c for c in contracts if c.remaining_tonnage <= 0]
    if status == "active":
        contracts = [c for c in contracts if c.is_active_contract]
    elif status == "completed":
        contracts = [c for c in contracts if c.is_completed]
    elif status == "pending":
        contracts = [c for c in contracts if c.pending_tonnage > 0]
    elif status == "review":
        contracts = [c for c in contracts if c.needs_review]
    return contracts


def seed_if_needed():
    if Company.query.count() == 0:
        for item in COMPANIES:
            db.session.add(Company(name=item["name"], short_name=item["short_name"]))
        db.session.commit()
    if User.query.count() == 0:
        admin = User(username="admin", full_name="مدیر سیستم", role="admin")
        admin.set_password("admin123")
        db.session.add(admin)
        employee = User(username="employee", full_name="کارشناس ثبت", role="employee")
        employee.set_password("employee123")
        db.session.add(employee)
        db.session.commit()
    if app.config.get("SEED_EXCEL", True) and Contract.query.count() == 0:
        name_to_company = {c.name: c for c in Company.query.all()}
        for company_name, path in default_excel_paths().items():
            company = name_to_company.get(company_name)
            if company and os.path.exists(path):
                try:
                    import_workbook(path, company.id)
                except Exception:
                    db.session.rollback()
                    app.logger.exception("Excel seed failed for %s", company_name)


@app.before_request
def ensure_db():
    if not getattr(app, "_db_ready", False):
        db.create_all()
        seed_if_needed()
        app._db_ready = True
    if "_csrf" not in session and request.endpoint != "static":
        session["_csrf"] = secrets.token_hex(16)
    if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
        if app.config.get("TESTING") and app.config.get("CSRF_DISABLE"):
            return None
        if _submitted_csrf_token() != session.get("_csrf"):
            if request.path.startswith("/api/"):
                return json_error("نشست نامعتبر است. صفحه را تازه کنید.", 403)
            flash("نشست نامعتبر است. صفحه را تازه کنید و دوباره تلاش کنید.", "error")
            return redirect(request.referrer or url_for("login"))
    return None


@app.context_processor
def inject_globals():
    user = current_user()
    companies = Company.query.order_by(Company.id).all() if getattr(app, "_db_ready", False) else []
    return {
        "current_user": user,
        "all_companies": companies,
        "statuses": STATUSES,
        "STATUS_FINAL": STATUS_FINAL,
        "STATUS_PENDING": STATUS_PENDING,
        "STATUS_CANCELLED": STATUS_CANCELLED,
        "csrf_token": csrf_token() if session else "",
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("dashboard"))
    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        user = User.query.filter_by(username=username, is_active=True).first()
        if user and user.check_password(password):
            session["user_id"] = user.id
            log_action("ورود", "user", user.id, f"ورود کاربر {user.username}")
            return redirect(url_for("dashboard"))
        error = "نام کاربری یا رمز عبور نادرست است."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    user = current_user()
    if user:
        log_action("خروج", "user", user.id)
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    return render_template("dashboard.html")


@app.route("/contracts")
@login_required
def contracts_page():
    return render_template("contracts.html")


@app.route("/contracts/new")
@login_required
def contract_new_page():
    return render_template("contract_form.html", contract=None)


@app.route("/contracts/<int:contract_id>")
@login_required
def contract_detail_page(contract_id):
    contract = active_contract_or_404(contract_id)
    return render_template("contract_detail.html", contract=contract)


@app.route("/contracts/<int:contract_id>/edit")
@login_required
def contract_edit_page(contract_id):
    contract = active_contract_or_404(contract_id)
    return render_template("contract_form.html", contract=contract)


@app.route("/declarations")
@login_required
def declarations_page():
    return render_template("declarations.html")


@app.route("/declarations/new")
@login_required
def declaration_new_page():
    contract_id = request.args.get("contract_id", type=int)
    contract = None
    if contract_id:
        found = db.session.get(Contract, contract_id)
        if found and not found.is_deleted:
            contract = found
    return render_template("declaration_form.html", declaration=None, preselected_contract=contract)


@app.route("/declarations/<int:declaration_id>/edit")
@login_required
def declaration_edit_page(declaration_id):
    declaration = active_declaration_or_404(declaration_id)
    return render_template("declaration_form.html", declaration=declaration, preselected_contract=declaration.contract)


@app.route("/reports")
@login_required
def reports_page():
    return render_template("reports.html")


@app.route("/users")
@admin_required
def users_page():
    users = User.query.order_by(User.id).all()
    return render_template("users.html", users=users)


@app.route("/settings")
@admin_required
def settings_page():
    logs = AuditLog.query.order_by(AuditLog.id.desc()).limit(50).all()
    backups = sorted(os.listdir(BACKUP_DIR), reverse=True) if os.path.isdir(BACKUP_DIR) else []
    return render_template("settings.html", logs=logs, backups=backups)


@app.route("/api/companies")
@login_required
def api_companies():
    companies = Company.query.order_by(Company.id).all()
    return jsonify({"ok": True, "items": [{"id": c.id, "name": c.name, "short_name": c.short_name} for c in companies]})


@app.route("/api/lookups")
@login_required
def api_lookups():
    company_id = request.args.get("company_id", type=int)
    query = Contract.query.filter(Contract.is_deleted.is_(False))
    if company_id:
        query = query.filter(Contract.company_id == company_id)
    contracts = query.all()
    products = sorted({c.product_type.strip() for c in contracts if c.product_type and c.product_type.strip()})
    locations = sorted({c.delivery_location.strip() for c in contracts if c.delivery_location and c.delivery_location.strip()})
    customers = sorted({c.customer_name.strip() for c in contracts if c.customer_name and c.customer_name.strip()})
    return jsonify({"ok": True, "products": products, "locations": locations, "customers": customers})


@app.route("/api/dashboard")
@login_required
def api_dashboard():
    contracts = apply_contract_filters(Contract.query)
    decls = []
    for contract in contracts:
        decls.extend(contract.active_declarations())
    payload = {
        "ok": True,
        "stats": {
            "contracts": len(contracts),
            "contract_quantity": round(sum(c.contract_quantity or 0 for c in contracts), 3),
            "declared_tonnage": round(sum(c.finalized_tonnage for c in contracts), 3),
            "remaining_tonnage": round(sum(c.remaining_tonnage for c in contracts), 3),
            "active_contracts": sum(1 for c in contracts if c.is_active_contract),
            "completed_contracts": sum(1 for c in contracts if c.is_completed),
            "cancelled_declarations": sum(1 for d in decls if d.status == STATUS_CANCELLED),
            "pending_declarations": sum(1 for d in decls if d.status == STATUS_PENDING),
            "review_items": sum(1 for c in contracts if c.needs_review),
        },
        "by_product": [],
        "recent_contracts": [c.to_dict() for c in contracts[:8]],
    }
    groups = {}
    for contract in contracts:
        key = (contract.product_type or "نامشخص", contract.delivery_location or "نامشخص")
        groups.setdefault(key, {"product": key[0], "location": key[1], "remaining": 0, "quantity": 0, "count": 0})
        groups[key]["remaining"] += contract.remaining_tonnage or 0
        groups[key]["quantity"] += contract.contract_quantity or 0
        groups[key]["count"] += 1
    payload["by_product"] = sorted(groups.values(), key=lambda x: x["remaining"], reverse=True)
    return jsonify(payload)


@app.route("/api/contracts")
@login_required
def api_contracts():
    contracts = apply_contract_filters(Contract.query)
    return jsonify({"ok": True, "items": [c.to_dict() for c in contracts], "count": len(contracts)})


@app.route("/api/contracts/<int:contract_id>")
@login_required
def api_contract_detail(contract_id):
    contract = active_contract_or_404(contract_id)
    return jsonify({"ok": True, "item": contract.to_dict(include_declarations=True)})


@app.route("/api/contracts", methods=["POST"])
@login_required
def api_create_contract():
    data = request.get_json(silent=True) or request.form
    try:
        company_id = int(data.get("company_id") or 0)
        company = db.session.get(Company, company_id)
        if not company:
            return json_error("شرکت را انتخاب کنید.")
        number = (data.get("contract_number") or "").strip()
        if not number:
            return json_error("شماره قرارداد الزامی است.")
        quantity = parse_float(data.get("contract_quantity"), "مقدار قرارداد")
        existing = Contract.query.filter_by(company_id=company_id, contract_number=number, is_deleted=False).first()
        confirm = str(data.get("confirm_duplicate") or "").lower() in {"1", "true", "yes"}
        if existing and not confirm:
            return jsonify(
                {
                    "ok": False,
                    "duplicate": True,
                    "error": "شماره قرارداد در این شرکت تکراری است. در صورت تأیید، قرارداد جدید ثبت می‌شود.",
                }
            ), 409
        contract = Contract(
            company_id=company_id,
            contract_number=number,
            contract_date=(data.get("contract_date") or "").strip(),
            customer_name=(data.get("customer_name") or "").strip(),
            contract_quantity=quantity,
            product_type=(data.get("product_type") or "").strip(),
            delivery_location=(data.get("delivery_location") or "").strip(),
            notes=(data.get("notes") or "").strip(),
        )
        if not contract.customer_name:
            return json_error("نام مشتری الزامی است.")
        db.session.add(contract)
        db.session.commit()
        log_action("ثبت قرارداد", "contract", contract.id, number)
        return jsonify({"ok": True, "item": contract.to_dict()}), 201
    except ValueError as exc:
        return json_error(str(exc))


@app.route("/api/contracts/<int:contract_id>", methods=["PUT"])
@login_required
def api_update_contract(contract_id):
    contract = active_contract_or_404(contract_id)
    data = request.get_json(silent=True) or request.form
    try:
        number = (data.get("contract_number") or "").strip()
        if not number:
            return json_error("شماره قرارداد الزامی است.")
        company_id = int(data.get("company_id") or contract.company_id)
        if not db.session.get(Company, company_id):
            return json_error("شرکت نامعتبر است.")
        existing = Contract.query.filter(
            Contract.company_id == company_id,
            Contract.contract_number == number,
            Contract.is_deleted.is_(False),
            Contract.id != contract.id,
        ).first()
        confirm = str(data.get("confirm_duplicate") or "").lower() in {"1", "true", "yes"}
        if existing and not confirm:
            return jsonify({"ok": False, "duplicate": True, "error": "شماره قرارداد در این شرکت تکراری است."}), 409
        quantity = parse_float(data.get("contract_quantity"), "مقدار قرارداد")
        remaining_after = quantity - contract.finalized_tonnage
        allow_over = str(data.get("allow_over_declaration") or "").lower() in {"1", "true", "yes"}
        if remaining_after < 0 and not (allow_over or contract.allow_over_declaration):
            return json_error("مقدار قرارداد از تناژ نهایی‌شده کمتر می‌شود. برای تأیید این مورد، گزینه مجوز تنظیم را فعال کنید.")
        contract.company_id = company_id
        contract.contract_number = number
        contract.contract_date = (data.get("contract_date") or "").strip()
        customer_name = (data.get("customer_name") or "").strip()
        if not customer_name:
            return json_error("نام مشتری الزامی است.")
        contract.customer_name = customer_name
        contract.contract_quantity = quantity
        contract.product_type = (data.get("product_type") or "").strip()
        contract.delivery_location = (data.get("delivery_location") or "").strip()
        contract.notes = (data.get("notes") or "").strip()
        contract.allow_over_declaration = allow_over or contract.allow_over_declaration
        contract.updated_at = datetime.utcnow()
        db.session.commit()
        log_action("ویرایش قرارداد", "contract", contract.id, number)
        return jsonify({"ok": True, "item": contract.to_dict(include_declarations=True)})
    except ValueError as exc:
        return json_error(str(exc))


@app.route("/api/contracts/<int:contract_id>", methods=["DELETE"])
@login_required
def api_delete_contract(contract_id):
    contract = active_contract_or_404(contract_id)
    if contract.active_declarations():
        return json_error("ابتدا کوتاژهای این قرارداد را حذف یا ابطال کنید.")
    contract.is_deleted = True
    db.session.commit()
    log_action("حذف قرارداد", "contract", contract.id, contract.contract_number)
    return jsonify({"ok": True})


@app.route("/api/declarations")
@login_required
def api_declarations():
    company_id = request.args.get("company_id", type=int)
    q = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()
    contract_id = request.args.get("contract_id", type=int)
    query = Declaration.query.join(Contract).filter(Declaration.is_deleted.is_(False), Contract.is_deleted.is_(False))
    if company_id:
        query = query.filter(Contract.company_id == company_id)
    if contract_id:
        query = query.filter(Declaration.contract_id == contract_id)
    if status:
        query = query.filter(Declaration.status == status)
    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Declaration.declaration_number.ilike(like),
                Declaration.dispatch_reference_number.ilike(like),
                Declaration.recipient.ilike(like),
                Contract.contract_number.ilike(like),
            )
        )
    items = query.order_by(Declaration.id.desc()).all()
    return jsonify({"ok": True, "items": [d.to_dict() for d in items], "count": len(items)})


def validate_declaration_tonnage(contract, tonnage, status, exclude_id=None, allow_over=False):
    if status != STATUS_FINAL:
        return
    current = sum(
        d.tonnage or 0
        for d in contract.active_declarations()
        if d.status == STATUS_FINAL and d.id != exclude_id
    )
    remaining = (contract.contract_quantity or 0) - current - tonnage
    if remaining < 0 and not (allow_over or contract.allow_over_declaration):
        raise ValueError(
            f"تناژ این کوتاژ از مانده قرارداد بیشتر است. مانده فعلی: {round((contract.contract_quantity or 0) - current, 3)}"
        )


@app.route("/api/declarations", methods=["POST"])
@login_required
def api_create_declaration():
    data = request.get_json(silent=True) or request.form
    try:
        contract_id = int(data.get("contract_id") or 0)
        contract = db.session.get(Contract, contract_id)
        if not contract or contract.is_deleted:
            return json_error("قرارداد معتبر انتخاب کنید.")
        number = (data.get("declaration_number") or "").strip()
        if not number:
            return json_error("شماره کوتاژ الزامی است.")
        tonnage = parse_float(data.get("tonnage"), "تناژ")
        status = (data.get("status") or STATUS_PENDING).strip()
        if status not in STATUSES:
            return json_error("وضعیت کوتاژ نامعتبر است.")
        allow_over = str(data.get("allow_over_declaration") or "").lower() in {"1", "true", "yes"}
        validate_declaration_tonnage(contract, tonnage, status, allow_over=allow_over)
        if allow_over:
            contract.allow_over_declaration = True
        declaration = Declaration(
            contract_id=contract.id,
            declaration_number=number,
            tonnage=tonnage,
            recipient=(data.get("recipient") or "").strip(),
            description=(data.get("description") or "").strip(),
            dispatch_reference_number=(data.get("dispatch_reference_number") or "").strip(),
            status=status,
        )
        db.session.add(declaration)
        db.session.commit()
        log_action("ثبت کوتاژ", "declaration", declaration.id, number)
        return jsonify({"ok": True, "item": declaration.to_dict(), "contract": contract.to_dict()}), 201
    except ValueError as exc:
        return json_error(str(exc))


@app.route("/api/declarations/<int:declaration_id>", methods=["PUT"])
@login_required
def api_update_declaration(declaration_id):
    declaration = active_declaration_or_404(declaration_id)
    data = request.get_json(silent=True) or request.form
    try:
        number = (data.get("declaration_number") or "").strip()
        if not number:
            return json_error("شماره کوتاژ الزامی است.")
        tonnage = parse_float(data.get("tonnage"), "تناژ")
        status = (data.get("status") or declaration.status).strip()
        if status not in STATUSES:
            return json_error("وضعیت کوتاژ نامعتبر است.")
        allow_over = str(data.get("allow_over_declaration") or "").lower() in {"1", "true", "yes"}
        validate_declaration_tonnage(declaration.contract, tonnage, status, exclude_id=declaration.id, allow_over=allow_over)
        if allow_over:
            declaration.contract.allow_over_declaration = True
        declaration.declaration_number = number
        declaration.tonnage = tonnage
        declaration.recipient = (data.get("recipient") or "").strip()
        declaration.description = (data.get("description") or "").strip()
        declaration.dispatch_reference_number = (data.get("dispatch_reference_number") or "").strip()
        declaration.status = status
        declaration.updated_at = datetime.utcnow()
        db.session.commit()
        log_action("ویرایش کوتاژ", "declaration", declaration.id, number)
        return jsonify({"ok": True, "item": declaration.to_dict(), "contract": declaration.contract.to_dict()})
    except ValueError as exc:
        return json_error(str(exc))


@app.route("/api/declarations/<int:declaration_id>/cancel", methods=["POST"])
@login_required
def api_cancel_declaration(declaration_id):
    declaration = active_declaration_or_404(declaration_id)
    declaration.status = STATUS_CANCELLED
    if not declaration.description:
        declaration.description = STATUS_CANCELLED
    declaration.updated_at = datetime.utcnow()
    db.session.commit()
    log_action("ابطال کوتاژ", "declaration", declaration.id, declaration.declaration_number)
    return jsonify({"ok": True, "item": declaration.to_dict(), "contract": declaration.contract.to_dict()})


@app.route("/api/declarations/<int:declaration_id>", methods=["DELETE"])
@login_required
def api_delete_declaration(declaration_id):
    declaration = active_declaration_or_404(declaration_id)
    declaration.is_deleted = True
    db.session.commit()
    log_action("حذف کوتاژ", "declaration", declaration.id, declaration.declaration_number)
    return jsonify({"ok": True, "contract": declaration.contract.to_dict()})


@app.route("/api/reports/export")
@login_required
def api_export():
    company_id = request.args.get("company_id", type=int)
    contracts = apply_contract_filters(Contract.query)
    if company_id:
        company = db.session.get(Company, company_id)
        if not company:
            return json_error("شرکت نامعتبر است.")
    elif contracts and len({c.company_id for c in contracts}) == 1:
        company = contracts[0].company
    else:
        company = type("CompanyObj", (), {"name": "همه شرکت‌ها", "short_name": "همه"})()
    title = f"گزارش کوتاژ و اظهار {company.name}"
    buf = export_contracts_workbook(company, contracts, title=title)
    filename = f"report-{company.short_name if hasattr(company, 'short_name') else 'all'}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/import", methods=["POST"])
@admin_required
def api_import():
    company_id = request.form.get("company_id", type=int)
    company = db.session.get(Company, company_id)
    if not company:
        flash("شرکت را انتخاب کنید.", "error")
        return redirect(url_for("settings_page"))
    file = request.files.get("file")
    if not file or not file.filename:
        flash("فایل اکسل را انتخاب کنید.", "error")
        return redirect(url_for("settings_page"))
    try:
        filename = safe_excel_filename(file.filename)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("settings_page"))
    path = os.path.join(UPLOAD_DIR, filename)
    file.save(path)
    try:
        result = import_workbook(path, company.id)
    except Exception:
        db.session.rollback()
        app.logger.exception("Excel import failed")
        flash("ورود فایل اکسل با خطا مواجه شد. فایل را بررسی کنید.", "error")
        return redirect(url_for("settings_page"))
    log_action("ورود اکسل", "company", company.id, f"{filename} / {result}")
    flash(f"ورود انجام شد: {result['contracts']} قرارداد و {result['declarations']} کوتاژ.", "success")
    return redirect(url_for("settings_page"))


@app.route("/api/backup", methods=["POST"])
@admin_required
def api_backup():
    src = os.path.join(INSTANCE_DIR, "kotaj.db")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"kotaj-{stamp}.db")
    shutil.copy2(src, dest)
    log_action("پشتیبان‌گیری", "database", None, dest)
    flash("نسخه پشتیبان با موفقیت ایجاد شد.", "success")
    return redirect(url_for("settings_page"))


@app.route("/api/backup/<name>")
@admin_required
def api_download_backup(name):
    safe = secure_filename(name)
    path = os.path.join(BACKUP_DIR, safe)
    if not os.path.exists(path):
        flash("فایل پشتیبان یافت نشد.", "error")
        return redirect(url_for("settings_page"))
    return send_file(path, as_attachment=True)


@app.route("/api/users", methods=["POST"])
@admin_required
def api_create_user():
    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""
    full_name = (request.form.get("full_name") or "").strip()
    role = request.form.get("role") or "employee"
    if not username or not password:
        flash("نام کاربری و رمز عبور الزامی است.", "error")
        return redirect(url_for("users_page"))
    if User.query.filter_by(username=username).first():
        flash("این نام کاربری قبلاً ثبت شده است.", "error")
        return redirect(url_for("users_page"))
    user = User(username=username, full_name=full_name, role=role if role in {"admin", "employee"} else "employee")
    user.set_password(password)
    db.session.add(user)
    db.session.commit()
    log_action("ایجاد کاربر", "user", user.id, username)
    flash("کاربر جدید ثبت شد.", "success")
    return redirect(url_for("users_page"))


@app.route("/api/users/<int:user_id>/toggle", methods=["POST"])
@admin_required
def api_toggle_user(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("کاربر یافت نشد.", "error")
        return redirect(url_for("users_page"))
    if user.id == session.get("user_id"):
        flash("امکان غیرفعال کردن حساب جاری وجود ندارد.", "error")
        return redirect(url_for("users_page"))
    user.is_active = not user.is_active
    db.session.commit()
    log_action("تغییر وضعیت کاربر", "user", user.id, str(user.is_active))
    return redirect(url_for("users_page"))


if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        seed_if_needed()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
