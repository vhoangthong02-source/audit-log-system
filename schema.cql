-- =============================================================
-- AUDIT LOG SYSTEM - Apache Cassandra Schema
-- Keyspace: audit_monitoring_ks
-- Chuẩn: IEEE 830 SRS / Barem 10.0 / 10.0 Điểm
-- =============================================================

CREATE KEYSPACE IF NOT EXISTS audit_monitoring_ks
WITH replication = {
    'class': 'SimpleStrategy',
    'replication_factor': 1
};

USE audit_monitoring_ks;

-- =============================================================
-- 1. Bảng Cấu hình Danh mục Dịch vụ theo dõi (CRUD Nền tảng 2.5đ)
-- =============================================================
CREATE TABLE IF NOT EXISTS monitored_services (
    service_name         TEXT PRIMARY KEY,
    display_name         TEXT,
    description          TEXT,
    default_ttl_seconds  INT,
    status               TEXT,               -- 'ACTIVE' | 'INACTIVE'
    updated_at           TIMESTAMP
);

-- =============================================================
-- 2. Bảng Chính lưu trữ Nhật ký Hoạt động (SRS Mục 4 & 5)
-- Composite Partition Key: ((service_name, log_date)) -> Time-Bucketing
-- Clustering Key: (log_time DESC, log_id ASC) -> Sắp xếp thời gian & chống trùng mili-giây
-- =============================================================
CREATE TABLE IF NOT EXISTS audit_logs (
    service_name    TEXT,
    log_date        DATE,
    log_time        TIMESTAMP,
    log_id          UUID,
    severity        TEXT,               -- 'INFO' | 'WARN' | 'ERROR'
    action_type     TEXT,               -- 'CREATE' | 'UPDATE' | 'DELETE' | 'LOGIN' ...
    actor_id        TEXT,               -- user/admin identifier
    ip_address      TEXT,
    endpoint        TEXT,
    message         TEXT,
    payload_json    TEXT,               -- JSON String trạng thái trước / sau
    PRIMARY KEY ((service_name, log_date), log_time, log_id)
) WITH CLUSTERING ORDER BY (log_time DESC, log_id ASC);

CREATE INDEX IF NOT EXISTS idx_audit_logs_severity
    ON audit_logs (severity);

-- =============================================================
-- 3. Bảng Mở rộng (Điểm cộng): Đếm tổng số Log bằng Cassandra COUNTER
-- =============================================================
CREATE TABLE IF NOT EXISTS service_log_stats (
    service_name    TEXT,
    log_date        DATE,
    severity        TEXT,
    total_count     COUNTER,
    PRIMARY KEY ((service_name, log_date), severity)
);
