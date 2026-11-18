import os
import sys
import tempfile

import pytest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app import app as flask_app
from models import db


@pytest.fixture()
def app():
    fd, db_path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    with flask_app.app_context():
        db.session.remove()
        db.engine.dispose()
    flask_app.config.update(
        TESTING=True,
        CSRF_DISABLE=True,
        SEED_EXCEL=False,
        SQLALCHEMY_DATABASE_URI="sqlite:///" + db_path.replace("\\", "/"),
    )
    flask_app._db_ready = False
    with flask_app.app_context():
        db.drop_all()
        db.create_all()
        from app import seed_if_needed

        seed_if_needed()
        flask_app._db_ready = True
    yield flask_app
    with flask_app.app_context():
        db.session.remove()
        db.drop_all()
        db.engine.dispose()
    try:
        os.remove(db_path)
    except OSError:
        pass


@pytest.fixture()
def client(app):
    return app.test_client()


@pytest.fixture()
def auth_client(client):
    client.post("/login", data={"username": "admin", "password": "admin123"})
    return client


@pytest.fixture()
def employee_client(client):
    client.post("/login", data={"username": "employee", "password": "employee123"})
    return client
