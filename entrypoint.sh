set -e

CASSANDRA_HOST="${CASSANDRA_HOST:-cassandra}"
CASSANDRA_PORT="${CASSANDRA_PORT:-9042}"
MAX_WAIT=120
ELAPSED=0

echo "⏳ Chờ Cassandra tại $CASSANDRA_HOST:$CASSANDRA_PORT..."

# Dùng Python socket để kiểm tra port 9042 mở chưa (không cần cqlsh)
until python -c "
import socket, sys
try:
    s = socket.create_connection(('$CASSANDRA_HOST', $CASSANDRA_PORT), timeout=3)
    s.close()
    sys.exit(0)
except Exception as e:
    sys.exit(1)
" 2>/dev/null; do
  if [ $ELAPSED -ge $MAX_WAIT ]; then
    echo "❌ Cassandra không phản hồi sau ${MAX_WAIT}s. Dừng."
    exit 1
  fi
  echo "   ... đang chờ TCP port (${ELAPSED}s / ${MAX_WAIT}s)"
  sleep 5
  ELAPSED=$((ELAPSED + 5))
done

echo "✅ TCP port mở! Chờ thêm 15s để Cassandra khởi tạo hoàn toàn..."
sleep 15

# Apply schema bằng Python (không cần cqlsh)
echo "📊 Áp dụng schema..."
python init_db.py
echo "✅ Schema + Seed hoàn thành!"

# Chạy FastAPI với --reload (hỗ trợ hot-reload khi mount volume)
echo "🚀 Khởi động FastAPI tại cổng 8000..."
exec python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
