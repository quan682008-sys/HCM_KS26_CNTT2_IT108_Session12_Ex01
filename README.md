# HCM_KS26_CNTT2_IT108_Session12_Ex01
##  Phần 1. Phân tích dữ liệu giao dịch PayKX

**Record (bản ghi):** Mỗi hàng trong bảng Excel đại diện cho một giao dịch ví điện tử.

**Field (trường dữ liệu):** Mỗi cột thể hiện một thuộc tính của giao dịch.

| Field | Ý nghĩa |
|---|---|
| Mã giao dịch | Mã định danh duy nhất |
| Mã người dùng | Xác định chủ giao dịch |
| Tên người dùng | Tra cứu bằng VLOOKUP |
| Ngày giao dịch | Thời điểm phát sinh |
| Loại giao dịch | Thanh toán hoặc nạp điện thoại |
| Ngân hàng | Đối tác xử lý giao dịch |
| Số tiền | Giá trị giao dịch (VNĐ) |
| Phí ngân hàng | Chi phí trả đối tác |
| Hoàn tiền | Số tiền ưu đãi khách hàng |
| Doanh thu | Khoản thu của PayKX |
| Lãi/Lỗ | Doanh thu trừ chi phí |

---

##  Phần 3. Kiểm thử lỗi #N/A

Khi VLOOKUP trả về #N/A, thực hiện 4 bước:

1. **Kiểm tra mã giao dịch:** Mã cần tìm có tồn tại trong bảng danh mục không?
2. **Làm sạch dữ liệu:** Dùng TRIM loại bỏ khoảng trắng thừa, kiểm tra kiểu dữ liệu và ký tự ẩn.
3. **Kiểm tra VLOOKUP:** Xác nhận đúng vùng tra cứu, cột trả về và sử dụng FALSE để tìm chính xác.
4. **Đối chiếu và chạy lại:** Sửa dữ liệu nguồn nếu sai, kiểm tra kết quả và sử dụng IFERROR để hiển thị cảnh báo phù hợp.

**Công thức áp dụng:**

```excel
=IFERROR(VLOOKUP(TRIM(A2),DanhMuc!$A$2:$B$25,2,FALSE),"CẦN ĐỐI SOÁT")
```
