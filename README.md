# Audit Log Monitoring System — Apache Cassandra & FastAPI

Hệ thống giám sát và lưu trữ nhật ký kiểm toán (Audit Log) tập trung dựa trên cơ sở dữ liệu phân tán **Apache Cassandra** và **FastAPI**.

---

## 🌟 Tính Năng Nổi Bật

- **Tối ưu hóa ghi tốc độ cao (Write-Heavy Optimized):** Sử dụng kiến trúc phân tán của Apache Cassandra với hiệu năng ghi vượt trội.
- **Cơ chế tự hủy bản ghi linh hoạt (Cell-Level TTL):** Quản lý vòng đời dữ liệu log với mệnh đề `USING TTL`, tự động dọn dẹp giải phóng đĩa.
- **Tra cứu đa chiều với Partition Key kép:** Thiết kế phân vùng `((service_name, log_date), log_time, log_id)` loại bỏ Full Cluster Scan.
- **Thống kê thời gian thực $O(1)$:** Sử dụng bảng `counter` phân tán để theo dõi chỉ số theo mức độ `INFO`, `WARN`, `ERROR` mà không dùng `COUNT(*)`.
- **Giao diện Dashboard trực quan:** Giao diện phẳng, hỗ trợ bộ lọc đa tiêu chí, đếm ngược TTL thời gian thực và trình giả lập sự kiện nghiệp vụ.
- **Đóng gói toàn diện với Docker Compose:** Khởi chạy toàn bộ hệ thống gồm Cassandra và FastAPI chỉ với 1 câu lệnh.

---

## 🏗️ Cấu Trúc Dự Án

```
├── db.py               # Kết nối Cassandra & Prepared Statements
├── main.py             # RESTful API Server (FastAPI)
├── schema.cql          # DDL định nghĩa Keyspace & Bảng dữ liệu
├── seed_data.py        # Kịch bản nạp 500 bản ghi dữ liệu mẫu
├── init_db.py          # Script tự động tạo bảng khi container khởi động
├── docker-compose.yml  # Điều phối cụm container Cassandra & App
├── Dockerfile          # Đóng gói ứng dụng backend Python 3.11-slim
├── entrypoint.sh       # Script khởi động container app
├── requirements.txt    # Danh sách thư viện phụ thuộc
├── static/
│   └── index.html      # Giao diện Web Dashboard (HTML/CSS/JS thuần)
└── tests/              # Bộ kiểm thử tự động
```

---

## 🚀 Hướng Dẫn Cài Đặt & Khởi Chạy

### Yêu Cầu
- [Docker](https://www.docker.com/) & Docker Compose

### Khởi Chạy Hệ Thống
```bash
# 1. Khởi động cụm dịch vụ
docker compose up -d

# 2. Kiểm tra trạng thái các container
docker compose ps
```

Sau khi các container ở trạng thái `healthy` / `Up`, truy cập hệ thống tại:
- **Web Dashboard:** `http://localhost:8000`
- **Swagger API Docs:** `http://localhost:8000/docs`
- **Cassandra CQL Port:** `9042`

---

## 🧪 Kiểm Thử

```bash
docker compose exec app pytest
```
