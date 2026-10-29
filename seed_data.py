import uuid
import json
import random
import logging
from datetime import datetime, timedelta, timezone
from db import get_session

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

SERVICES = [
    {
        "service_name":        "Auth",
        "display_name":        "Authentication Service",
        "description":         "Xac thuc danh tinh, cap phat JWT va kiem soat truy cap OAuth2",
        "default_ttl_seconds": 604800,
        "status":              "ACTIVE",
    },
    {
        "service_name":        "Payment",
        "display_name":        "Payment Gateway Service",
        "description":         "Xu ly thanh toan the, QR-Pay, vi dien tu va doi soat ngan hang",
        "default_ttl_seconds": 2592000,
        "status":              "ACTIVE",
    },
    {
        "service_name":        "Order",
        "display_name":        "Order Management Service",
        "description":         "Tiep nhan gio hang, dieu phoi van chuyen va quan ly don hang",
        "default_ttl_seconds": 1209600,
        "status":              "ACTIVE",
    },
    {
        "service_name":        "User",
        "display_name":        "User Profile Service",
        "description":         "Quan ly thong tin tai khoan, ho so khach hang va dinh danh eKYC",
        "default_ttl_seconds": 604800,
        "status":              "ACTIVE",
    },
]

SERVICE_TEMPLATES = {
    "Auth": {
        "actions": ["LOGIN", "LOGOUT", "TOKEN_REFRESH", "REVOKE_ROLE"],
        "endpoints": ["/api/v1/auth/login", "/api/v1/auth/token/refresh", "/api/v1/auth/revoke", "/api/v1/auth/logout"],
        "messages": {
            "INFO":  "Dang nhap OAuth2 thanh cong, cap moi cap JWT Access/Refresh Token.",
            "WARN":  "Phat hien dang nhap sai mat khau 3 lan lien tiep tu dia chi IP la.",
            "ERROR": "Khoa tam thoi tai khoan do dau hieu tan cong Brute-force; Token gia mao chu ky.",
        },
        "payload_builder": lambda sev, act: {
            "auth_type": "OAuth2_PKCE",
            "client_id": "web-portal-v2",
            "attempt_count": 1 if sev == "INFO" else (3 if sev == "WARN" else 10),
            "device": random.choice(["Chrome / Windows 11", "Safari / macOS Sonoma", "Firefox / Ubuntu 24.04"]),
            "token_id": str(uuid.uuid4())[:8],
        }
    },
    "Payment": {
        "actions": ["CREATE", "UPDATE", "REFUND", "WEBHOOK_SYNC"],
        "endpoints": ["/api/v1/payments/charge", "/api/v1/payments/refund", "/api/v1/payments/webhook", "/api/v1/payments/status"],
        "messages": {
            "INFO":  "Khoi tao va khau tru giao dich thanh toan QR-Pay thanh cong.",
            "WARN":  "Cong thanh toan ngan hang phan hoi cham vuot nguong canh bao 4.500ms.",
            "ERROR": "Giao dich that bai do so du khong du (INSUFFICIENT_FUNDS) hoac loi chu ky doi soat.",
        },
        "payload_builder": lambda sev, act: {
            "transaction_id": f"TXN-{random.randint(10000000, 99999999)}",
            "amount": round(random.uniform(50_000, 15_000_000), 2),
            "currency": "VND",
            "gateway": random.choice(["VNPay-QR", "VietQR-Napas", "Momo-Enterprise", "ZaloPay"]),
            "latency_ms": random.randint(120, 450) if sev == "INFO" else (random.randint(4501, 8000) if sev == "WARN" else 0),
            "error_code": None if sev == "INFO" else ("ERR_LATENCY_TIMEOUT" if sev == "WARN" else "INSUFFICIENT_FUNDS"),
        }
    },
    "Order": {
        "actions": ["CREATE", "UPDATE", "DELETE", "DISPATCH"],
        "endpoints": ["/api/v1/orders/checkout", "/api/v1/orders/cancel", "/api/v1/orders/dispatch", "/api/v1/orders/update-status"],
        "messages": {
            "INFO":  "Tao moi don hang, cap nhat trang thai ban giao don vi van chuyen.",
            "WARN":  "So luong ton kho cua ma san pham trong don cham nguong canh bao toi thieu.",
            "ERROR": "Huy don hang khan cap do loi khoa ban ghi kho hang (Inventory Deadlock).",
        },
        "payload_builder": lambda sev, act: {
            "order_code": f"ORD-2026-{random.randint(10000, 99999)}",
            "total_items": random.randint(1, 8),
            "shipping_carrier": random.choice(["GiaoHangTietKiem", "ViettelPost", "ShopeeXpress", "J&T Express"]),
            "sku": random.choice(["MACBOOK-M4-PRO", "IPHONE-16-PROMAX", "DELL-XPS-15", "THINKPAD-X1-CARBON"]),
            "before_status": "PROCESSING" if act in ["UPDATE", "DELETE"] else "NEW",
            "after_status": "CANCELLED" if sev == "ERROR" else ("DISPATCHED" if act == "DISPATCH" else "CONFIRMED"),
        }
    },
    "User": {
        "actions": ["CREATE", "UPDATE", "DELETE", "KYC_VERIFY"],
        "endpoints": ["/api/v1/users/profile", "/api/v1/users/kyc", "/api/v1/users/delete", "/api/v1/users/change-phone"],
        "messages": {
            "INFO":  "Nguoi dung cap nhat ho so ca nhan, thay doi dia chi nhan hang mac dinh.",
            "WARN":  "Yeu cau doi so dien thoai bao mat nhung het han ma xac thuc OTP.",
            "ERROR": "Tu choi thao tac xoa tai khoan do van con cong no hoac don hang dang giao.",
        },
        "payload_builder": lambda sev, act: {
            "account_status": "ACTIVE" if sev != "ERROR" else "LOCKED_PENDING_AUDIT",
            "kyc_tier": random.choice(["Tier_1_Basic", "Tier_2_CitizenID", "Tier_3_Biometric"]),
            "outstanding_balance": 0 if sev != "ERROR" else round(random.uniform(500_000, 3_500_000), 2),
            "field_modified": random.choice(["shipping_address", "phone_number", "avatar_url"]),
        }
    },
}

ACTORS = [
    "admin_super", "admin_khovanhanh", "operator_hanoi", "operator_hcm",
    "service_cron_job", "system_reconciler", "user_nguyenvana", "user_lethib"
]

IPS = [
    "192.168.10.24", "192.168.10.88", "10.0.12.45", "10.0.15.102",
    "172.16.0.5", "172.16.1.18", "203.113.152.1", "118.70.188.42"
]


def seed_services(session):
    prep = session.prepare("""
        INSERT INTO monitored_services
            (service_name, display_name, description, default_ttl_seconds, status, updated_at)
        VALUES (?, ?, ?, ?, ?, toTimestamp(now()))
    """)
    for svc in SERVICES:
        session.execute(prep, (
            svc["service_name"],
            svc["display_name"],
            svc["description"],
            svc["default_ttl_seconds"],
            svc["status"],
        ))


def seed_logs(session, total_per_service: int = 125):
    prep_insert = session.prepare("""
        INSERT INTO audit_logs
            (service_name, log_date, log_time, log_id,
             severity, action_type, actor_id, ip_address, endpoint, message, payload_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        USING TTL ?
    """)

    prep_counter = session.prepare("""
        UPDATE service_log_stats
        SET total_count = total_count + 1
        WHERE service_name = ? AND log_date = ? AND severity = ?
    """)

    now = datetime.now(timezone.utc)

    for svc in SERVICES:
        svc_name = svc["service_name"]
        svc_ttl  = svc["default_ttl_seconds"]
        tmpl     = SERVICE_TEMPLATES[svc_name]

        plan = [("INFO", 80), ("WARN", 25), ("ERROR", 20)]
        records_list = []
        for sev, count in plan:
            records_list.extend([sev] * count)
        random.shuffle(records_list)

        for i, severity in enumerate(records_list):
            action_type  = random.choice(tmpl["actions"])
            endpoint     = random.choice(tmpl["endpoints"])
            message      = tmpl["messages"][severity]
            payload_data = tmpl["payload_builder"](severity, action_type)
            payload_json = json.dumps(payload_data, ensure_ascii=False)

            if i < 20:
                delta_minutes = random.randint(5, 480)
                log_time = now - timedelta(minutes=delta_minutes)
            else:
                delta_days  = random.randint(1, 6)
                delta_hours = random.randint(0, 23)
                delta_mins  = random.randint(0, 59)
                log_time = now - timedelta(days=delta_days, hours=delta_hours, minutes=delta_mins)

            log_date   = log_time.date()
            log_id     = uuid.uuid4()
            actor_id   = random.choice(ACTORS)
            ip_address = random.choice(IPS)

            session.execute(prep_insert, (
                svc_name, log_date, log_time, log_id,
                severity, action_type, actor_id, ip_address, endpoint,
                message, payload_json, svc_ttl
            ))

            session.execute(prep_counter, (svc_name, log_date, severity))


def main():
    session = get_session()
    seed_services(session)
    seed_logs(session)


if __name__ == "__main__":
    main()
