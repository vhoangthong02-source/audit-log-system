import os
import sys
import uuid
from datetime import datetime, date, timezone
from unittest.mock import MagicMock, patch

import pytest
from fastapi.testclient import TestClient

# Import app
import main
from main import app, ServiceCreate, ServiceUpdate, LogIngestRequest

client = TestClient(app, raise_server_exceptions=False)

# Fixture for mock DB
@pytest.fixture(autouse=True)
def mock_db():
    with patch("main.db") as mock_db_module:
        mock_session = MagicMock()
        mock_db_module.get_session.return_value = mock_session
        mock_db_module.get_prepared.side_effect = lambda k: f"PREPARED_{k}"
        mock_db_module.build_insert_log_cql.side_effect = lambda ttl: f"INSERT INTO audit_logs ... USING TTL {ttl}"
        yield mock_session, mock_db_module


def test_health_check_success(mock_db):
    mock_session, _ = mock_db
    mock_row = {"release_version": "4.1.0"}
    mock_result = MagicMock()
    mock_result.one.return_value = mock_row
    mock_session.execute.return_value = mock_result

    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "ok"
    assert data["cassandra"] == "connected"
    assert data["release_version"] == "4.1.0"


def test_list_services(mock_db):
    mock_session, _ = mock_db
    mock_session.execute.return_value = [
        {"service_name": "Payment", "display_name": "Payment Service"},
        {"service_name": "Auth", "display_name": "Auth Service"},
    ]

    response = client.get("/api/v1/services")
    assert response.status_code == 200
    data = response.json()["data"]
    # Should be sorted by service_name
    assert [d["service_name"] for d in data] == ["Auth", "Payment"]


def test_get_service_found(mock_db):
    mock_session, _ = mock_db
    mock_result = MagicMock()
    mock_result.one.return_value = {"service_name": "Auth", "status": "ACTIVE"}
    mock_session.execute.return_value = mock_result

    response = client.get("/api/v1/services/Auth")
    assert response.status_code == 200
    assert response.json()["data"]["service_name"] == "Auth"


def test_get_service_not_found(mock_db):
    mock_session, _ = mock_db
    mock_result = MagicMock()
    mock_result.one.return_value = None
    mock_session.execute.return_value = mock_result

    response = client.get("/api/v1/services/NonExistent")
    assert response.status_code == 404


def test_create_service_success(mock_db):
    mock_session, _ = mock_db
    # First query check existing -> None
    mock_result = MagicMock()
    mock_result.one.return_value = None
    mock_session.execute.return_value = mock_result

    payload = {
        "service_name": "Notification",
        "display_name": "Notification Service",
        "description": "Email and SMS notifications",
        "default_ttl_seconds": 86400,
        "status": "ACTIVE"
    }
    response = client.post("/api/v1/services", json=payload)
    assert response.status_code == 201
    assert response.json()["status"] == "SUCCESS"


def test_create_service_already_exists(mock_db):
    mock_session, _ = mock_db
    mock_result = MagicMock()
    mock_result.one.return_value = {"service_name": "Notification"}
    mock_session.execute.return_value = mock_result

    payload = {
        "service_name": "Notification",
        "display_name": "Notification Service",
        "description": "Email and SMS notifications",
        "default_ttl_seconds": 86400,
        "status": "ACTIVE"
    }
    response = client.post("/api/v1/services", json=payload)
    assert response.status_code == 409


def test_ingest_log_success(mock_db):
    mock_session, mock_db_module = mock_db
    # Service check
    mock_result = MagicMock()
    mock_result.one.return_value = {
        "service_name": "Auth",
        "status": "ACTIVE",
        "default_ttl_seconds": 604800
    }
    mock_session.execute.return_value = mock_result

    payload = {
        "service_name": "Auth",
        "severity": "INFO",
        "action_type": "LOGIN",
        "actor_id": "admin",
        "ip_address": "127.0.0.1",
        "endpoint": "/api/v1/auth/login",
        "message": "User login success",
        "payload_json": '{"user": "test"}'
    }
    response = client.post("/api/v1/logs", json=payload)
    assert response.status_code == 201
    data = response.json()
    assert data["status"] == "SUCCESS"
    assert data["applied_ttl_seconds"] == 604800
    assert "partition_key" in data
    assert "log_id" in data


def test_ingest_log_inactive_service(mock_db):
    mock_session, _ = mock_db
    mock_result = MagicMock()
    mock_result.one.return_value = {
        "service_name": "Auth",
        "status": "INACTIVE",
        "default_ttl_seconds": 604800
    }
    mock_session.execute.return_value = mock_result

    payload = {
        "service_name": "Auth",
        "severity": "INFO",
        "message": "Test"
    }
    response = client.post("/api/v1/logs", json=payload)
    assert response.status_code == 400
    assert "INACTIVE" in response.json()["detail"]


def test_ingest_log_service_not_found(mock_db):
    mock_session, _ = mock_db
    mock_result = MagicMock()
    mock_result.one.return_value = None
    mock_session.execute.return_value = mock_result

    payload = {
        "service_name": "GhostService",
        "severity": "INFO",
        "message": "Test"
    }
    response = client.post("/api/v1/logs", json=payload)
    assert response.status_code == 404


def test_search_logs(mock_db):
    mock_session, _ = mock_db
    now = datetime.now(timezone.utc)
    mock_session.execute.return_value = [
        {
            "service_name": "Auth",
            "log_date": date.today(),
            "log_time": now,
            "log_id": uuid.uuid4(),
            "severity": "INFO",
            "action_type": "LOGIN",
            "actor_id": "user1",
            "ip_address": "127.0.0.1",
            "endpoint": "/api/v1/auth",
            "message": "Test log",
            "payload_json": "{}",
            "remaining_ttl": 604000
        }
    ]

    response = client.get("/api/v1/logs/search?service_name=Auth")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "SUCCESS"
    assert data["total"] == 1
    assert data["data"][0]["remaining_ttl"] == 604000


def test_get_stats(mock_db):
    mock_session, _ = mock_db
    mock_session.execute.return_value = [
        {"service_name": "Auth", "log_date": date.today(), "severity": "INFO", "total_count": 10},
        {"service_name": "Auth", "log_date": date.today(), "severity": "ERROR", "total_count": 2},
        {"service_name": "Payment", "log_date": date.today(), "severity": "WARN", "total_count": 5},
    ]

    response = client.get("/api/v1/stats")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "SUCCESS"
    assert data["total_logs"] == 17
    assert data["info_count"] == 10
    assert data["error_count"] == 2
    assert data["warn_count"] == 5
    assert "Auth" in data["by_service"]
    assert data["by_service"]["Auth"]["total"] == 12
