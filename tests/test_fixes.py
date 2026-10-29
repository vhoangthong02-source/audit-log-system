import uuid
from datetime import datetime, date, timezone, timedelta
from unittest.mock import MagicMock, patch
import pytest
from fastapi.testclient import TestClient

import main
from main import app, parse_iso_datetime

client = TestClient(app, raise_server_exceptions=False)


def test_fix_severity_filtered_at_db():
    """Xác nhận lỗi 1 đã sửa: Lọc severity được thực thi trực tiếp tại DB qua prepared statement"""
    with patch("main.db") as mock_db:
        mock_session = MagicMock()
        mock_db.get_session.return_value = mock_session
        mock_db.get_prepared.side_effect = lambda k: f"PREPARED_{k}"

        now = datetime.now(timezone.utc)
        error_rows = [
            {"service_name": "Auth", "log_date": date.today(), "log_time": now,
             "log_id": uuid.uuid4(), "severity": "ERROR", "action_type": "LOGIN", "actor_id": "u",
             "ip_address": "1.1.1.1", "endpoint": "/a", "message": "msg", "payload_json": "{}", "remaining_ttl": 100}
        ]
        mock_session.execute.return_value = error_rows

        response = client.get("/api/v1/logs/search?service_name=Auth&severity=ERROR&limit=5")
        data = response.json()

        assert response.status_code == 200
        assert data["status"] == "SUCCESS"
        assert len(data["data"]) == 1
        assert "AND severity = 'ERROR'" in data["executed_cql"]

        called_statement = mock_session.execute.call_args[0][0]
        assert called_statement == "PREPARED_select_logs_by_date_and_severity"
        called_params = mock_session.execute.call_args[0][1]
        assert called_params == ("Auth", date.today(), "ERROR", 5)


def test_fix_timezone_conversion():
    """Xác nhận lỗi 2 đã sửa: Chuyển đổi múi giờ ISO an toàn chính xác sang UTC"""
    vn_time_str = "2026-10-26T10:00:00+07:00"
    dt = parse_iso_datetime(vn_time_str)

    assert dt.hour == 3
    assert dt.tzinfo == timezone.utc


def test_fix_invalid_severity_rejected():
    """Xác nhận lỗi Pydantic validation: Reject mức độ nghiêm trọng không hợp lệ"""
    payload = {
        "service_name": "Auth",
        "severity": "CRITICAL_INVALID",
        "message": "Test"
    }
    response = client.post("/api/v1/logs", json=payload)
    assert response.status_code == 422  


def test_fix_insert_log_uses_prepared_statement():
    """Xác nhận lỗi 3 đã sửa: Ghi log sử dụng Prepared Statement thay vì format CQL thô"""
    with patch("main.db") as mock_db:
        mock_session = MagicMock()
        mock_db.get_session.return_value = mock_session
        mock_db.get_prepared.side_effect = lambda k: f"PREPARED_{k}"

        mock_svc = MagicMock()
        mock_svc.one.return_value = {
            "service_name": "Auth",
            "status": "ACTIVE",
            "default_ttl_seconds": 604800
        }
        mock_session.execute.return_value = mock_svc

        payload = {
            "service_name": "Auth",
            "severity": "INFO",
            "message": "Testing prepared insert",
            "ttl_seconds": 120
        }
        response = client.post("/api/v1/logs", json=payload)
        assert response.status_code == 201

        # Cuộc gọi execute thứ 2 là insert log
        insert_call = mock_session.execute.call_args_list[1]
        called_statement = insert_call[0][0]
        called_params = insert_call[0][1]

        assert called_statement == "PREPARED_insert_log"
        assert called_params[11] == 120  
