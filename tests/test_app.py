import io
import os
import shutil
import tempfile

from openpyxl import load_workbook

from models import STATUS_CANCELLED, STATUS_FINAL, STATUS_PENDING, Company, Contract, Declaration, db
from utils.excel_io import export_contracts_workbook, import_workbook

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")


def login(client, username="admin", password="admin123"):
    return client.post("/login", data={"username": username, "password": password}, follow_redirects=True)


def create_contract(client, company_id=1, number="C-100", quantity=10000, **extra):
    payload = {
        "company_id": company_id,
        "contract_number": number,
        "contract_date": "1405/01/15",
        "customer_name": extra.pop("customer_name", "مشتری آزمایشی"),
        "contract_quantity": quantity,
        "product_type": extra.pop("product_type", "سیمان"),
        "delivery_location": extra.pop("delivery_location", "EXW-میل 78"),
    }
    payload.update(extra)
    return client.post("/api/contracts", json=payload)


def create_declaration(client, contract_id, number, tonnage, status=STATUS_FINAL, **extra):
    payload = {
        "contract_id": contract_id,
        "declaration_number": number,
        "tonnage": tonnage,
        "recipient": extra.pop("recipient", "آقای تست"),
        "dispatch_reference_number": extra.pop("dispatch_reference_number", "H-1"),
        "description": extra.pop("description", ""),
        "status": status,
    }
    payload.update(extra)
    return client.post("/api/declarations", json=payload)


def test_application_startup_and_pages(auth_client):
    for path in ["/", "/contracts", "/contracts/new", "/declarations", "/declarations/new", "/reports", "/users", "/settings"]:
        res = auth_client.get(path)
        assert res.status_code == 200, path


def test_unauthenticated_pages_redirect(client):
    res = client.get("/")
    assert res.status_code in (302, 401)
    res = client.get("/api/contracts")
    assert res.status_code == 401


def test_database_initialization(app):
    with app.app_context():
        companies = Company.query.all()
        assert len(companies) == 2
        names = {c.name for c in companies}
        assert "شرکت تولیدی معدنی آتشکده خراسان" in names
        assert "شرکت گچ خراسان" in names
        assert db.session.query(db.func.count()).select_from(Contract).scalar() == 0


def test_contract_create_view_edit_search(auth_client):
    res = create_contract(auth_client, number="1405999001", quantity=1500.5)
    assert res.status_code == 201
    item = res.get_json()["item"]
    cid = item["id"]
    assert item["contract_number"] == "1405999001"
    assert item["contract_date"] == "1405/01/15"
    assert item["remaining_tonnage"] == 1500.5

    res = auth_client.get(f"/contracts/{cid}")
    assert res.status_code == 200
    res = auth_client.get(f"/api/contracts/{cid}")
    assert res.get_json()["item"]["customer_name"] == "مشتری آزمایشی"

    res = auth_client.put(
        f"/api/contracts/{cid}",
        json={
            "company_id": 1,
            "contract_number": "1405999001",
            "contract_date": "1405/02/01",
            "customer_name": "مشتری ویرایش‌شده",
            "contract_quantity": 1600,
            "product_type": "سیمان پرتلند",
            "delivery_location": "EXW-دوغارون",
        },
    )
    assert res.status_code == 200
    assert res.get_json()["item"]["customer_name"] == "مشتری ویرایش‌شده"
    assert res.get_json()["item"]["contract_quantity"] == 1600

    res = auth_client.get("/api/contracts?q=1405999001")
    assert res.get_json()["count"] == 1
    res = auth_client.get("/api/contracts?q=ویرایش")
    assert res.get_json()["count"] == 1


def test_duplicate_and_invalid_contract(auth_client):
    assert create_contract(auth_client, number="DUP-1").status_code == 201
    res = create_contract(auth_client, number="DUP-1")
    assert res.status_code == 409
    assert res.get_json()["duplicate"] is True
    res = create_contract(auth_client, number="DUP-1", confirm_duplicate=True)
    assert res.status_code == 201
    res = create_contract(auth_client, number="BAD", quantity=-10)
    assert res.status_code == 400
    res = auth_client.post("/api/contracts", json={"company_id": 1, "contract_number": "X"})
    assert res.status_code == 400


def test_multiple_declarations_and_persistence(auth_client, app):
    cid = create_contract(auth_client, number="DEC-1", quantity=10000).get_json()["item"]["id"]
    create_declaration(auth_client, cid, "D1", 1000)
    create_declaration(auth_client, cid, "D2", 2000, recipient="آقای ب", dispatch_reference_number="H-2")
    create_declaration(auth_client, cid, "D3", 500, status=STATUS_PENDING, description="در حال بررسی")
    detail = auth_client.get(f"/api/contracts/{cid}").get_json()["item"]
    assert detail["declaration_count"] == 3
    assert detail["finalized_tonnage"] == 3000
    assert detail["pending_tonnage"] == 500
    assert detail["remaining_tonnage"] == 7000

    with app.app_context():
        db.session.remove()
        saved = db.session.get(Contract, cid)
        assert saved is not None
        assert len(saved.active_declarations()) == 3
        assert saved.remaining_tonnage == 7000


def test_edit_and_cancel_declaration(auth_client):
    cid = create_contract(auth_client, number="DEC-2", quantity=10000).get_json()["item"]["id"]
    d1 = create_declaration(auth_client, cid, "E1", 1000).get_json()["item"]["id"]
    d2 = create_declaration(auth_client, cid, "E2", 2000).get_json()["item"]["id"]
    res = auth_client.put(
        f"/api/declarations/{d2}",
        json={"declaration_number": "E2", "tonnage": 2500, "status": STATUS_FINAL, "recipient": "ب"},
    )
    assert res.status_code == 200
    assert res.get_json()["contract"]["remaining_tonnage"] == 6500
    res = auth_client.post(f"/api/declarations/{d1}/cancel")
    assert res.status_code == 200
    contract = res.get_json()["contract"]
    assert contract["remaining_tonnage"] == 7500
    assert contract["cancelled_tonnage"] == 1000
    assert contract["finalized_tonnage"] == 2500


def test_remaining_tonnage_scenarios(auth_client):
    cid = create_contract(auth_client, number="CALC-A", quantity=10000).get_json()["item"]["id"]
    create_declaration(auth_client, cid, "A1", 1000)
    create_declaration(auth_client, cid, "A2", 2000)
    item = auth_client.get(f"/api/contracts/{cid}").get_json()["item"]
    assert item["remaining_tonnage"] == 7000

    create_declaration(auth_client, cid, "A3", 500, status=STATUS_PENDING)
    item = auth_client.get(f"/api/contracts/{cid}").get_json()["item"]
    assert item["remaining_tonnage"] == 7000
    assert item["pending_tonnage"] == 500

    d1 = next(d for d in item["declarations"] if d["declaration_number"] == "A1")
    auth_client.post(f"/api/declarations/{d1['id']}/cancel")
    item = auth_client.get(f"/api/contracts/{cid}").get_json()["item"]
    assert item["remaining_tonnage"] == 8000

    cid2 = create_contract(auth_client, number="CALC-D", quantity=10000).get_json()["item"]["id"]
    create_declaration(auth_client, cid2, "D1", 10000)
    item = auth_client.get(f"/api/contracts/{cid2}").get_json()["item"]
    assert item["remaining_tonnage"] == 0
    assert item["is_completed"] is True


def test_zero_decimal_negative_and_over_declaration(auth_client):
    cid = create_contract(auth_client, number="EDGE-1", quantity=0).get_json()["item"]["id"]
    assert auth_client.get(f"/api/contracts/{cid}").get_json()["item"]["remaining_tonnage"] == 0
    cid = create_contract(auth_client, number="EDGE-2", quantity=10.25).get_json()["item"]["id"]
    res = create_declaration(auth_client, cid, "DEC", 1.5)
    assert res.status_code == 201
    assert res.get_json()["contract"]["remaining_tonnage"] == 8.75
    res = create_declaration(auth_client, cid, "NEG", -1)
    assert res.status_code == 400
    res = create_declaration(auth_client, cid, "OVER", 20)
    assert res.status_code == 400
    res = create_declaration(auth_client, cid, "OVER2", 20, allow_over_declaration=True)
    assert res.status_code == 201
    assert res.get_json()["contract"]["remaining_tonnage"] < 0


def test_company_separation(auth_client):
    a = create_contract(auth_client, company_id=1, number="SAME-1", customer_name="آتشکده").get_json()["item"]
    g = create_contract(auth_client, company_id=2, number="SAME-1", customer_name="گچ خراسان").get_json()["item"]
    assert a["company_id"] != g["company_id"]
    atash = auth_client.get("/api/contracts?company_id=1").get_json()["items"]
    gach = auth_client.get("/api/contracts?company_id=2").get_json()["items"]
    assert all(item["company_id"] == 1 for item in atash)
    assert all(item["company_id"] == 2 for item in gach)
    dash_a = auth_client.get("/api/dashboard?company_id=1").get_json()["stats"]
    dash_g = auth_client.get("/api/dashboard?company_id=2").get_json()["stats"]
    assert dash_a["contracts"] == 1
    assert dash_g["contracts"] == 1


def test_search_and_filters(auth_client):
    create_contract(
        auth_client,
        company_id=1,
        number="SRCH-11",
        quantity=100,
        customer_name="شرکت نمونه یک",
        product_type="کیسه",
        delivery_location="میلک",
        contract_date="1405/01/10",
    )
    create_contract(
        auth_client,
        company_id=2,
        number="SRCH-22",
        quantity=200,
        customer_name="شرکت نمونه دو",
        product_type="فله",
        delivery_location="دوغارون",
        contract_date="1405/02/10",
    )
    cid = auth_client.get("/api/contracts?q=SRCH-11").get_json()["items"][0]["id"]
    create_declaration(auth_client, cid, "KOTAJ-77", 10, dispatch_reference_number="HAV-9")
    assert auth_client.get("/api/contracts?q=نمونه یک").get_json()["count"] == 1
    assert auth_client.get("/api/contracts?product=کیسه").get_json()["count"] == 1
    assert auth_client.get("/api/contracts?location=دوغارون").get_json()["count"] == 1
    assert auth_client.get("/api/contracts?company_id=2&q=SRCH").get_json()["count"] == 1
    assert auth_client.get("/api/contracts?q=وجودندارد").get_json()["count"] == 0
    assert auth_client.get("/api/declarations?q=KOTAJ-77").get_json()["count"] == 1
    assert auth_client.get("/api/declarations?q=HAV-9").get_json()["count"] == 1
    assert auth_client.get("/api/contracts?date_from=1405/02/01").get_json()["count"] == 1


def test_excel_export_preserves_text_and_totals(auth_client, app):
    cid = create_contract(auth_client, number="1405000012", quantity=1000).get_json()["item"]["id"]
    create_declaration(auth_client, cid, "000123", 250, dispatch_reference_number="14050019")
    res = auth_client.get("/api/reports/export?company_id=1")
    assert res.status_code == 200
    assert "spreadsheetml" in res.headers["Content-Type"]
    wb = load_workbook(io.BytesIO(res.data))
    ws = wb.active
    assert ws["B3"].value == "1405000012"
    assert ws["B3"].number_format == "@"
    assert ws["C3"].value == "1405/01/15"
    assert ws["I3"].value == 750
    with app.app_context():
        company = db.session.get(Company, 1)
        contracts = Contract.query.filter_by(company_id=1, is_deleted=False).all()
        buf = export_contracts_workbook(company, contracts)
        exported = load_workbook(buf)
        sheet = exported.active
        numbers = [sheet.cell(row, 2).value for row in range(3, 20) if sheet.cell(row, 2).value]
        assert "1405000012" in numbers


def test_excel_import_from_copies(app):
    with app.app_context():
        company = Company.query.filter_by(name="شرکت گچ خراسان").one()
        src = os.path.join(DATA_DIR, "gach_khorasan.xlsx")
        tmp = os.path.join(tempfile.gettempdir(), "gach_copy.xlsx")
        shutil.copy2(src, tmp)
        result = import_workbook(tmp, company.id)
        assert result["contracts"] >= 1
        assert result["declarations"] >= 1
        contract = Contract.query.filter_by(company_id=company.id, contract_number="1405031999").first()
        assert contract is not None
        cancelled = [d for d in contract.active_declarations() if d.status == STATUS_CANCELLED]
        assert len(cancelled) == 2
        assert contract.cancelled_tonnage == 1000
        assert contract.finalized_tonnage == 4250
        assert contract.remaining_tonnage == 750
        os.remove(tmp)


def test_employee_cannot_access_admin_pages(employee_client):
    res = employee_client.get("/users", follow_redirects=False)
    assert res.status_code in (302, 403)
    res = employee_client.get("/settings", follow_redirects=False)
    assert res.status_code in (302, 403)


def test_soft_delete_does_not_corrupt_related_records(auth_client, app):
    cid = create_contract(auth_client, number="DEL-1", quantity=100).get_json()["item"]["id"]
    did = create_declaration(auth_client, cid, "DX", 10).get_json()["item"]["id"]
    res = auth_client.delete(f"/api/contracts/{cid}")
    assert res.status_code == 400
    auth_client.delete(f"/api/declarations/{did}")
    res = auth_client.delete(f"/api/contracts/{cid}")
    assert res.status_code == 200
    assert auth_client.get(f"/api/contracts/{cid}").status_code == 404
    with app.app_context():
        assert db.session.get(Contract, cid).is_deleted is True
        assert db.session.get(Declaration, did).is_deleted is True


def test_csrf_required_when_enabled(app, client):
    login(client)
    app.config["CSRF_DISABLE"] = False
    res = client.post(
        "/api/contracts",
        json={
            "company_id": 1,
            "contract_number": "NO-CSRF",
            "customer_name": "x",
            "contract_quantity": 1,
        },
    )
    assert res.status_code == 403
