# Trinh Buyer – Tạo file Shipment Sample (SS)

Công cụ web chạy hoàn toàn trên trình duyệt (không gửi dữ liệu đi đâu):

1. Thả file **PO follow-up** (.xlsx).
2. Chọn tìm theo **Số PO** hoặc **Số Invoice**, nhập danh sách → *Tìm dòng*.
3. Nhập Batch no#, kiểm tra số lượng, thông tin chung (Buyer, ngày gửi RM, ngày NCC gửi, AWB).
4. Bấm **Tạo file SS** → tải file .xlsx đúng cột sheet "2026" của file SS follow-up.

## Cột SS lấy từ đâu

| SS | PO follow-up |
|---|---|
| PO | PO No. / PO SCAF |
| SUPPLIER | Supplier |
| Item | Article Supplier (thiếu thì RM Name) |
| Code color | Supplier color code |
| Color | Color |
| Customer | Customer Code |
| Story | Story |
| BU name | BU name |
| quantity/M | Order Quantity (tìm theo INV: Qty_Inv) |

Công cụ tự nhận mọi sheet có cột `PO No.` + `Supplier` (PO follow-up- tracking, SHIPPED, DATA SCAX…), bỏ qua sheet PIVOT,
gộp dòng theo PO + Item + Color, ưu tiên dữ liệu SAP ở sheet *tracking*, và đọc lại các dòng bị dán lệch cột ở sheet SHIPPED (gắn nhãn "lệch cột").

Mã nguồn: `index.html` (giao diện) + `core.js` (đọc file, ghép dữ liệu, xuất SS). Thư viện: SheetJS 0.18.5.
