import os
import logging
from cassandra.cluster import Cluster
from cassandra.policies import DCAwareRoundRobinPolicy
from cassandra.query import dict_factory

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

CASSANDRA_HOST = os.getenv("CASSANDRA_HOST", "cassandra")
CASSANDRA_PORT = int(os.getenv("CASSANDRA_PORT", "9042"))
CASSANDRA_DC   = os.getenv("CASSANDRA_DC", "dc1")


def apply_schema(session_no_ks):
    with open("schema.cql", encoding="utf-8") as f:
        content = f.read()

    statements = [s.strip() for s in content.split(";") if s.strip()]
    ok = 0
    for stmt in statements:
        lines = [l for l in stmt.splitlines() if not l.strip().startswith("--")]
        stmt_clean = "\n".join(lines).strip()
        if not stmt_clean:
            continue
        try:
            session_no_ks.execute(stmt_clean)
            ok += 1
        except Exception as e:
            logger.warning(f"Skip stmt: {e} | CQL: {stmt_clean[:60]}...")
    logger.info(f"Schema applied ({ok} statements)")


def main():
    cluster = Cluster(
        contact_points=[CASSANDRA_HOST],
        port=CASSANDRA_PORT,
        load_balancing_policy=DCAwareRoundRobinPolicy(local_dc=CASSANDRA_DC),
        protocol_version=5,
        connect_timeout=15,
    )
    session = cluster.connect()
    session.row_factory = dict_factory

    apply_schema(session)

    session.set_keyspace("audit_monitoring_ks")

    from seed_data import seed_services, seed_logs
    seed_services(session)
    seed_logs(session)

    cluster.shutdown()
    logger.info("Init DB done")


if __name__ == "__main__":
    main()
