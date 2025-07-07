### 🧪 Technical Field

Hệ thống phát minh này thuộc lĩnh vực bảo mật ứng dụng Android, đặc biệt là kỹ thuật phát hiện phần mềm độc hại dựa trên hành vi runtime.

- Áp dụng cho các hệ thống phân tích hành vi tự động.
- Có thể triển khai như một công cụ sandbox hoặc tích hợp vào hệ thống MDM.

![](../_assets/images/sample_diagram.png)

---

### 🔍 Kỹ thuật cốt lõi

- Mô phỏng tương tác người dùng qua script điều khiển.
- Thu thập hành vi ứng dụng khi tương tác như click, scroll, nhập liệu.
- Phân tích các mẫu hành vi qua thống kê thời gian phản hồi, số lượng API gọi.

Hệ thống này đặc biệt phù hợp để phân tích các APK chứa đoạn mã độc được kích hoạt bằng các trigger ẩn (delay, điều kiện môi trường...).
