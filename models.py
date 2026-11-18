from datetime import datetime

from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import check_password_hash, generate_password_hash

db = SQLAlchemy()

STATUS_FINAL = "کوتاژ نهایی شد"
STATUS_CANCELLED = "کوتاژ ابطال شد"
STATUS_PENDING = "در انتظار تکمیل"

STATUSES = [STATUS_FINAL, STATUS_PENDING, STATUS_CANCELLED]


class Company(db.Model):
    __tablename__ = "companies"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), unique=True, nullable=False)
    short_name = db.Column(db.String(100), nullable=False)

    contracts = db.relationship("Contract", backref="company", lazy="dynamic")


class User(db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120), nullable=False, default="")
    role = db.Column(db.String(20), nullable=False, default="employee")
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    @property
    def is_admin(self):
        return self.role == "admin"


class Contract(db.Model):
    __tablename__ = "contracts"

    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey("companies.id"), nullable=False)
    contract_number = db.Column(db.String(64), nullable=False)
    contract_date = db.Column(db.String(20), nullable=False, default="")
    customer_name = db.Column(db.String(255), nullable=False, default="")
    contract_quantity = db.Column(db.Float, nullable=False, default=0)
    product_type = db.Column(db.String(255), nullable=False, default="")
    delivery_location = db.Column(db.String(255), nullable=False, default="")
    notes = db.Column(db.Text, nullable=False, default="")
    allow_over_declaration = db.Column(db.Boolean, nullable=False, default=False)
    needs_review = db.Column(db.Boolean, nullable=False, default=False)
    review_note = db.Column(db.Text, nullable=False, default="")
    is_deleted = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    declarations = db.relationship(
        "Declaration",
        backref="contract",
        lazy="dynamic",
        cascade="all, delete-orphan",
    )

    def active_declarations(self):
        return [d for d in self.declarations if not d.is_deleted]

    def tonnage_by_status(self, status):
        return sum(d.tonnage or 0 for d in self.active_declarations() if d.status == status)

    @property
    def finalized_tonnage(self):
        return self.tonnage_by_status(STATUS_FINAL)

    @property
    def pending_tonnage(self):
        return self.tonnage_by_status(STATUS_PENDING)

    @property
    def cancelled_tonnage(self):
        return self.tonnage_by_status(STATUS_CANCELLED)

    @property
    def valid_declared_tonnage(self):
        return self.finalized_tonnage

    @property
    def remaining_tonnage(self):
        return round((self.contract_quantity or 0) - self.valid_declared_tonnage, 3)

    @property
    def is_completed(self):
        return self.remaining_tonnage <= 0 and self.finalized_tonnage > 0

    @property
    def is_active_contract(self):
        return self.remaining_tonnage > 0

    @property
    def declaration_numbers_text(self):
        nums = [d.declaration_number for d in self.active_declarations() if d.declaration_number]
        return "/".join(nums)

    def to_dict(self, include_declarations=False):
        data = {
            "id": self.id,
            "company_id": self.company_id,
            "company_name": self.company.name if self.company else "",
            "contract_number": self.contract_number,
            "contract_date": self.contract_date,
            "customer_name": self.customer_name,
            "contract_quantity": self.contract_quantity,
            "product_type": self.product_type,
            "delivery_location": self.delivery_location,
            "notes": self.notes,
            "allow_over_declaration": self.allow_over_declaration,
            "needs_review": self.needs_review,
            "review_note": self.review_note,
            "finalized_tonnage": self.finalized_tonnage,
            "pending_tonnage": self.pending_tonnage,
            "cancelled_tonnage": self.cancelled_tonnage,
            "remaining_tonnage": self.remaining_tonnage,
            "is_completed": self.is_completed,
            "declaration_numbers": self.declaration_numbers_text,
            "declaration_count": len(self.active_declarations()),
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
        }
        if include_declarations:
            data["declarations"] = [d.to_dict() for d in self.active_declarations()]
        return data


class Declaration(db.Model):
    __tablename__ = "declarations"

    id = db.Column(db.Integer, primary_key=True)
    contract_id = db.Column(db.Integer, db.ForeignKey("contracts.id"), nullable=False)
    declaration_number = db.Column(db.String(64), nullable=False)
    tonnage = db.Column(db.Float, nullable=False, default=0)
    recipient = db.Column(db.String(255), nullable=False, default="")
    description = db.Column(db.Text, nullable=False, default="")
    dispatch_reference_number = db.Column(db.String(128), nullable=False, default="")
    status = db.Column(db.String(64), nullable=False, default=STATUS_PENDING)
    needs_review = db.Column(db.Boolean, nullable=False, default=False)
    review_note = db.Column(db.Text, nullable=False, default="")
    is_deleted = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        company = self.contract.company if self.contract else None
        return {
            "id": self.id,
            "contract_id": self.contract_id,
            "contract_number": self.contract.contract_number if self.contract else "",
            "company_id": self.contract.company_id if self.contract else None,
            "company_name": company.name if company else "",
            "declaration_number": self.declaration_number,
            "tonnage": self.tonnage,
            "recipient": self.recipient,
            "description": self.description,
            "dispatch_reference_number": self.dispatch_reference_number,
            "status": self.status,
            "needs_review": self.needs_review,
            "review_note": self.review_note,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
        }


class AuditLog(db.Model):
    __tablename__ = "audit_logs"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    action = db.Column(db.String(64), nullable=False)
    record_type = db.Column(db.String(64), nullable=False, default="")
    record_id = db.Column(db.Integer, nullable=True)
    details = db.Column(db.Text, nullable=False, default="")
    timestamp = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    user = db.relationship("User", backref="audit_logs")
