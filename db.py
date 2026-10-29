import os
import time
import logging
from typing import Optional

from cassandra.cluster import Cluster
from cassandra.auth import PlainTextAuthProvider
from cassandra.policies import DCAwareRoundRobinPolicy
from cassandra.query import dict_factory

logger = logging.getLogger(__name__)

CASSANDRA_HOST = os.getenv("CASSANDRA_HOST", "localhost")
CASSANDRA_PORT = int(os.getenv("CASSANDRA_PORT", "9042"))
CASSANDRA_DC   = os.getenv("CASSANDRA_DC", "dc1")
CASSANDRA_USER = os.getenv("CASSANDRA_USER", "")
CASSANDRA_PASS = os.getenv("CASSANDRA_PASS", "")
KEYSPACE       = "audit_monitoring_ks"

_cluster: Optional[Cluster] = None
_session = None
_prepared: dict = {}


def get_session():
    global _cluster, _session

    if _session is not None:
        return _session

    auth = (
        PlainTextAuthProvider(username=CASSANDRA_USER, password=CASSANDRA_PASS)
        if CASSANDRA_USER else None
    )

    max_retries = 20
    for attempt in range(1, max_retries + 1):
        try:
            _cluster = Cluster(
                contact_points=[CASSANDRA_HOST],
                port=CASSANDRA_PORT,
                auth_provider=auth,
                load_balancing_policy=DCAwareRoundRobinPolicy(local_dc=CASSANDRA_DC),
                protocol_version=5,
                connect_timeout=10,
            )
            _session = _cluster.connect(KEYSPACE)
            _session.row_factory = dict_factory
            break
        except Exception as e:
            logger.warning(f"Connecting Cassandra attempt {attempt}/{max_retries}: {e}")
            if attempt == max_retries:
                raise RuntimeError(f"Cannot connect Cassandra after {max_retries} retries") from e
            time.sleep(3)

    _prepare_statements()
    return _session


def _prepare_statements():
    sess = _session

    _prepared["select_all_services"] = sess.prepare("""
        SELECT * FROM monitored_services
    """)

    _prepared["insert_service"] = sess.prepare("""
        INSERT INTO monitored_services
            (service_name, display_name, description, default_ttl_seconds, status, updated_at)
        VALUES (?, ?, ?, ?, ?, toTimestamp(now()))
    """)

    _prepared["select_service"] = sess.prepare("""
        SELECT * FROM monitored_services WHERE service_name = ?
    """)

    _prepared["update_service"] = sess.prepare("""
        UPDATE monitored_services SET
            display_name        = ?,
            description         = ?,
            default_ttl_seconds = ?,
            status              = ?,
            updated_at          = toTimestamp(now())
        WHERE service_name = ?
    """)

    _prepared["delete_service"] = sess.prepare("""
        DELETE FROM monitored_services WHERE service_name = ?
    """)

    _prepared["insert_log"] = sess.prepare("""
        INSERT INTO audit_logs
            (service_name, log_date, log_time, log_id,
             severity, action_type, actor_id, ip_address, endpoint, message, payload_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        USING TTL ?
    """)

    _prepared["insert_log_no_ttl"] = sess.prepare("""
        INSERT INTO audit_logs
            (service_name, log_date, log_time, log_id,
             severity, action_type, actor_id, ip_address, endpoint, message, payload_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """)

    _prepared["select_logs_by_date"] = sess.prepare("""
        SELECT service_name, log_date, log_time, log_id,
               severity, action_type, actor_id, ip_address, endpoint,
               message, payload_json, TTL(message) AS remaining_ttl
        FROM audit_logs
        WHERE service_name = ? AND log_date = ?
        LIMIT ?
    """)

    _prepared["select_logs_by_date_and_severity"] = sess.prepare("""
        SELECT service_name, log_date, log_time, log_id,
               severity, action_type, actor_id, ip_address, endpoint,
               message, payload_json, TTL(message) AS remaining_ttl
        FROM audit_logs
        WHERE service_name = ? AND log_date = ? AND severity = ?
        LIMIT ?
    """)

    _prepared["update_counter"] = sess.prepare("""
        UPDATE service_log_stats
        SET total_count = total_count + 1
        WHERE service_name = ? AND log_date = ? AND severity = ?
    """)

    _prepared["select_all_stats"] = sess.prepare("""
        SELECT service_name, log_date, severity, total_count FROM service_log_stats
    """)


def get_prepared(key: str):
    return _prepared.get(key)


def build_insert_log_cql(ttl: int) -> str:
    return f"""
        INSERT INTO audit_logs
            (service_name, log_date, log_time, log_id,
             severity, action_type, actor_id, ip_address, endpoint, message, payload_json)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        USING TTL {ttl}
    """


def close():
    global _cluster, _session
    if _cluster:
        _cluster.shutdown()
        _cluster = None
        _session = None
