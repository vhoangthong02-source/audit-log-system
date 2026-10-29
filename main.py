import uuid
import logging
from datetime import date, datetime, timezone
from typing import Optional, List, Dict, Any
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException, Query, status
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, field_validator

import db
from seed_data import seed_services, seed_logs

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    db.get_session()
    yield
    db.close()


app = FastAPI(
    title="He Thong Giam Sat Audit Log - Apache Cassandra",
    version="2.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def root():
    with open("static/index.html", encoding="utf-8") as f:
        return f.read()


class ServiceCreate(BaseModel):
    service_name:        str = Field(...)
    display_name:        str = Field(...)
    description:         str = Field("")
    default_ttl_seconds: int = Field(604800, ge=30, le=25920000)
    status:              str = Field("ACTIVE")

    @field_validator("status", mode="before")
    @classmethod
    def validate_status(cls, v: Any) -> str:
        if isinstance(v, str):
            v_upper = v.strip().upper()
            if v_upper in {"ACTIVE", "INACTIVE"}:
                return v_upper
        raise ValueError("status phai la ACTIVE hoac INACTIVE")


class ServiceUpdate(BaseModel):
    display_name:        Optional[str] = None
    description:         Optional[str] = None
    default_ttl_seconds: Optional[int] = Field(None, ge=30, le=25920000)
    status:              Optional[str] = None

    @field_validator("status", mode="before")
    @classmethod
    def validate_status(cls, v: Any) -> Optional[str]:
        if v is None:
            return None
        if isinstance(v, str):
            v_upper = v.strip().upper()
            if v_upper in {"ACTIVE", "INACTIVE"}:
                return v_upper
        raise ValueError("status phai la ACTIVE hoac INACTIVE")


class LogIngestRequest(BaseModel):
    service_name: str           = Field(...)
    severity:     str           = Field(...)
    action_type:  Optional[str] = Field(None)
    action:       Optional[str] = Field(None)
    actor_id:     Optional[str] = Field(None)
    user_id:      Optional[str] = Field(None)
    ip_address:   Optional[str] = Field("127.0.0.1")
    endpoint:     Optional[str] = Field("/api/v1/action")
    message:      str           = Field(...)
    payload_json: Optional[str] = Field(None)
    payload:      Optional[str] = Field(None)
    ttl_seconds:  Optional[int] = Field(None, ge=5, le=25920000)

    @field_validator("severity", mode="before")
    @classmethod
    def validate_severity(cls, v: Any) -> str:
        if isinstance(v, str):
            v_upper = v.strip().upper()
            if v_upper in {"INFO", "WARN", "ERROR"}:
                return v_upper
        raise ValueError("severity phai la INFO, WARN hoac ERROR")


def parse_iso_datetime(dt_str: str) -> datetime:
    dt = datetime.fromisoformat(dt_str)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


@app.get("/health", tags=["System"])
def health_check():
    try:
        session = db.get_session()
        row = session.execute("SELECT release_version FROM system.local").one()
        return {
            "status": "ok",
            "cassandra": "connected",
            "release_version": row["release_version"] if row else "unknown",
            "timestamp": datetime.now(timezone.utc).isoformat()
        }
    except Exception as e:
        raise HTTPException(status.HTTP_503_SERVICE_UNAVAILABLE, detail=str(e))


@app.get("/api/v1/services", tags=["Services"])
def list_services():
    session = db.get_session()
    prepared = db.get_prepared("select_all_services")
    rows = list(session.execute(prepared) if prepared else session.execute("SELECT * FROM monitored_services"))
    rows.sort(key=lambda x: x.get("service_name", ""))
    return {"data": rows}


@app.get("/api/v1/services/{service_name}", tags=["Services"])
def get_service(service_name: str):
    session = db.get_session()
    row = session.execute(db.get_prepared("select_service"), (service_name,)).one()
    if not row:
        raise HTTPException(404, detail=f"Khong tim thay dich vu '{service_name}'")
    return {"data": row}


@app.post("/api/v1/services", status_code=status.HTTP_201_CREATED, tags=["Services"])
def create_service(body: ServiceCreate):
    session = db.get_session()
    existing = session.execute(db.get_prepared("select_service"), (body.service_name,)).one()
    if existing:
        raise HTTPException(409, detail=f"Dich vu '{body.service_name}' da ton tai")

    session.execute(db.get_prepared("insert_service"), (
        body.service_name,
        body.display_name,
        body.description,
        body.default_ttl_seconds,
        body.status.upper(),
    ))
    return {
        "status": "SUCCESS",
        "message": f"Dich vu '{body.service_name}' da duoc them moi",
        "data": body.model_dump()
    }


@app.put("/api/v1/services/{service_name}", tags=["Services"])
def update_service(service_name: str, body: ServiceUpdate):
    session = db.get_session()
    existing = session.execute(db.get_prepared("select_service"), (service_name,)).one()
    if not existing:
        raise HTTPException(404, detail=f"Khong tim thay dich vu '{service_name}'")

    new_display = body.display_name if body.display_name is not None else existing["display_name"]
    new_desc    = body.description if body.description is not None else existing["description"]
    new_ttl     = body.default_ttl_seconds if body.default_ttl_seconds is not None else existing["default_ttl_seconds"]
    new_status  = body.status.upper() if body.status is not None else existing["status"]

    session.execute(db.get_prepared("update_service"), (
        new_display, new_desc, new_ttl, new_status, service_name
    ))
    return {
        "status": "SUCCESS",
        "message": f"Dich vu '{service_name}' da duoc cap nhat",
        "data": {
            "service_name": service_name,
            "display_name": new_display,
            "default_ttl_seconds": new_ttl,
            "status": new_status,
        }
    }


@app.delete("/api/v1/services/{service_name}", tags=["Services"])
def delete_service(service_name: str):
    session = db.get_session()
    existing = session.execute(db.get_prepared("select_service"), (service_name,)).one()
    if not existing:
        raise HTTPException(404, detail=f"Khong tim thay dich vu '{service_name}'")

    session.execute(db.get_prepared("delete_service"), (service_name,))
    return {
        "status": "SUCCESS",
        "message": f"Dich vu '{service_name}' da duoc xoa"
    }


@app.post("/api/v1/logs", status_code=status.HTTP_201_CREATED, tags=["Logs"])
def ingest_log(body: LogIngestRequest):
    session = db.get_session()

    svc = session.execute(db.get_prepared("select_service"), (body.service_name,)).one()
    if not svc:
        raise HTTPException(404, detail=f"Dich vu '{body.service_name}' chua duoc dang ky")
    if svc.get("status") == "INACTIVE":
        raise HTTPException(400, detail=f"Dich vu '{body.service_name}' dang INACTIVE")

    ttl = body.ttl_seconds if body.ttl_seconds is not None else svc.get("default_ttl_seconds", 86400)

    now          = datetime.now(timezone.utc)
    log_date     = now.date()
    log_id       = uuid.uuid4()
    severity     = body.severity.upper()
    action_type  = (body.action_type or body.action or "EXECUTE").upper()
    actor_id     = body.actor_id or body.user_id or "system"
    payload_json = body.payload_json or body.payload or "{}"
    endpoint     = body.endpoint or "/api/v1/action"
    ip_address   = body.ip_address or "127.0.0.1"

    if ttl and ttl > 0:
        session.execute(db.get_prepared("insert_log"), (
            body.service_name, log_date, now, log_id,
            severity, action_type, actor_id, ip_address, endpoint,
            body.message, payload_json, ttl
        ))
        executed_cql = (
            f"INSERT INTO audit_logs (service_name, log_date, log_time, log_id, severity, action_type, actor_id, ip_address, endpoint, message, payload_json) "
            f"VALUES ('{body.service_name}', '{log_date}', '{now.isoformat()}', {log_id}, '{severity}', '{action_type}', '{actor_id}', '{ip_address}', '{endpoint}', '{body.message}', '{payload_json}') "
            f"USING TTL {ttl};"
        )
    else:
        session.execute(db.get_prepared("insert_log_no_ttl"), (
            body.service_name, log_date, now, log_id,
            severity, action_type, actor_id, ip_address, endpoint,
            body.message, payload_json
        ))
        executed_cql = (
            f"INSERT INTO audit_logs (service_name, log_date, log_time, log_id, severity, action_type, actor_id, ip_address, endpoint, message, payload_json) "
            f"VALUES ('{body.service_name}', '{log_date}', '{now.isoformat()}', {log_id}, '{severity}', '{action_type}', '{actor_id}', '{ip_address}', '{endpoint}', '{body.message}', '{payload_json}');"
        )

    session.execute(db.get_prepared("update_counter"), (
        body.service_name, log_date, severity
    ))

    return {
        "status": "SUCCESS",
        "partition_key": f"({body.service_name}, {log_date})",
        "clustering_key": f"{now.isoformat()} | {log_id}",
        "applied_ttl_seconds": ttl,
        "executed_cql": executed_cql,
        "log_id": str(log_id),
        "log_time": now.isoformat() + "Z",
    }


@app.get("/api/v1/logs/search", tags=["Logs"])
def search_logs(
    service_name: str           = Query(...),
    log_date:     Optional[str] = Query(None),
    start_time:   Optional[str] = Query(None),
    end_time:     Optional[str] = Query(None),
    severity:     Optional[str] = Query(None),
    limit:        int           = Query(50, ge=1, le=500),
):
    session = db.get_session()
    filter_date = date.fromisoformat(log_date) if log_date else date.today()

    if severity:
        sev_target = severity.strip().upper()
        raw_rows = list(session.execute(
            db.get_prepared("select_logs_by_date_and_severity"),
            (service_name, filter_date, sev_target, limit),
        ))
        executed_cql = (
            f"SELECT service_name, log_date, log_time, log_id,\n"
            f"       severity, action_type, actor_id, ip_address, endpoint,\n"
            f"       message, payload_json, TTL(message) AS remaining_ttl\n"
            f"FROM audit_logs\n"
            f"WHERE service_name = '{service_name}'\n"
            f"  AND log_date = '{filter_date}'\n"
            f"  AND severity = '{sev_target}'\n"
            f"LIMIT {limit};"
        )
    else:
        raw_rows = list(session.execute(
            db.get_prepared("select_logs_by_date"),
            (service_name, filter_date, limit),
        ))
        executed_cql = (
            f"SELECT service_name, log_date, log_time, log_id,\n"
            f"       severity, action_type, actor_id, ip_address, endpoint,\n"
            f"       message, payload_json, TTL(message) AS remaining_ttl\n"
            f"FROM audit_logs\n"
            f"WHERE service_name = '{service_name}'\n"
            f"  AND log_date = '{filter_date}'\n"
            f"LIMIT {limit};"
        )

    filtered = raw_rows
    if start_time:
        try:
            dt_start = parse_iso_datetime(start_time)
            filtered = [r for r in filtered if r.get("log_time") and r["log_time"] >= dt_start]
        except Exception as e:
            logger.warning(f"start_time parse error: {e}")

    if end_time:
        try:
            dt_end = parse_iso_datetime(end_time)
            filtered = [r for r in filtered if r.get("log_time") and r["log_time"] <= dt_end]
        except Exception as e:
            logger.warning(f"end_time parse error: {e}")

    data = []
    for r in filtered:
        data.append({
            "service_name":  r["service_name"],
            "log_date":      str(r["log_date"]),
            "log_time":      (r["log_time"].isoformat() + "Z") if r.get("log_time") else None,
            "log_id":        str(r["log_id"]),
            "severity":      r.get("severity", "INFO"),
            "action_type":   r.get("action_type") or "ACTION",
            "actor_id":      r.get("actor_id") or "unknown",
            "ip_address":    r.get("ip_address") or "",
            "endpoint":      r.get("endpoint") or "",
            "message":       r.get("message") or "",
            "payload_json":  r.get("payload_json") or "{}",
            "remaining_ttl": int(r["remaining_ttl"]) if r.get("remaining_ttl") is not None else None,
        })

    return {
        "status": "SUCCESS",
        "total": len(data),
        "service_name": service_name,
        "log_date": str(filter_date),
        "executed_cql": executed_cql,
        "data": data,
    }


@app.get("/api/v1/stats", tags=["Stats"])
def get_stats():
    session = db.get_session()
    prepared = db.get_prepared("select_all_stats")
    rows = list(session.execute(prepared) if prepared else session.execute("SELECT service_name, log_date, severity, total_count FROM service_log_stats"))

    total_logs = 0
    info_count = 0
    warn_count = 0
    error_count = 0
    by_service: Dict[str, Dict[str, int]] = {}

    for r in rows:
        count = int(r.get("total_count") or 0)
        sev   = (r.get("severity") or "").upper()
        svc   = r.get("service_name") or "Unknown"

        total_logs += count
        if sev == "INFO":
            info_count += count
        elif sev == "WARN":
            warn_count += count
        elif sev == "ERROR":
            error_count += count

        if svc not in by_service:
            by_service[svc] = {"total": 0, "INFO": 0, "WARN": 0, "ERROR": 0}
        by_service[svc]["total"] += count
        if sev in by_service[svc]:
            by_service[svc][sev] += count

    return {
        "status": "SUCCESS",
        "total_logs": total_logs,
        "info_count": info_count,
        "warn_count": warn_count,
        "error_count": error_count,
        "by_service": by_service,
        "raw": [
            {
                "service_name": r["service_name"],
                "log_date": str(r["log_date"]),
                "severity": r["severity"],
                "total_count": int(r.get("total_count") or 0),
            }
            for r in rows
        ]
    }


@app.post("/api/v1/seed", tags=["System"])
def trigger_seed():
    session = db.get_session()
    seed_services(session)
    seed_logs(session)
    return {
        "status": "SUCCESS",
        "message": "Da nap thanh cong 4 dich vu va 500 ban ghi Audit Log!"
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
