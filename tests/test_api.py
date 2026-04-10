"""Basic health and convert endpoint tests (no docker required)."""
import sys
import os
import pytest

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)


def test_root():
    response = client.get("/")
    assert response.status_code == 200
    data = response.json()
    assert data["name"] == "SGML Pipeline API"
    assert data["version"] == "0.0.1"


def test_health():
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] in ("healthy", "degraded")
    assert "pipeline" in data


def test_convert_rejects_non_docx():
    """Uploading a non-.docx file should return 422."""
    response = client.post(
        "/convert",
        files={"file": ("test.txt", b"hello", "text/plain")},
    )
    assert response.status_code == 422


def test_convert_rejects_empty_file():
    """Uploading an empty .docx file should return 422."""
    response = client.post(
        "/convert",
        files={"file": ("test.docx", b"", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )
    assert response.status_code == 422
