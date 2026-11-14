# Web bảng chỉnh định rơ-le HPDQ

## Mở web
- Giải nén ZIP và mở thư mục relay-hpdq trong VS Code.
- Nếu dùng Live Server, mở index.html bằng Live Server. Hoặc mở Terminal ngay trong thư mục có index.html, chạy `python -m http.server 8000`, rồi vào http://localhost:8000/.
- Chọn HPDQ1 hoặc HPDQ2, mở nhóm NM.CĐ và bảng cần xem. Tìm theo một phần tên hoặc mã tủ. Excel xuất ra gồm toàn bộ bảng được chọn, kể cả khi màn hình đang lọc.

## Cập nhật nguồn
- Thay `du-lieu/HPDQ1.xlsx` hoặc `du-lieu/HPDQ2.xlsx`, giữ nguyên tên và tải lại trang.
- Nếu chỉnh file trên Google Drive, tải bản Excel mới về và thay vào thư mục này. Nếu chạy trên máy chủ, thay bản trên máy chủ.
- Nên giữ tên sheet và cách bố trí tiêu đề. Nếu thay đổi cấu trúc, kiểm tra lại tiêu đề và đơn vị trong bảng trước khi sử dụng.
- Bản HTML xem nhanh chứa bản chụp của hai file. Dùng mã nguồn trong ZIP để thay file về sau.

## Chỗ cần biết khi đọc code
- `index.html` bố cục, `style.css` giao diện, `excel.js` đọc và xuất XLSX, `app.js` cấu hình danh mục, tách bảng, lọc và hiển thị.
- `confirmedUnits` trong app.js khai báo sheet thuộc NM.CĐ nào. Bản thử gán theo thông tin có trong file. Các sheet còn lại nằm dưới “Chưa xác định NM.CĐ”.
- Quan hệ trạm cha/con trong bản thử suy ra từ thứ tự và tiêu đề trong sheet. Các mục 110 kV được giữ riêng. Những trường hợp không có tiêu đề phù hợp xuất hiện dưới “Chưa xác định trạm 35 kV”.
- Excel xuất dữ liệu bảng theo giá trị, với một sheet cho mỗi bảng khi xuất nhóm. Công thức và định dạng trình bày gốc nằm trong hai file nguồn.
- Một số ô thông số đã mang định dạng ngày trong file nguồn. Đối chiếu file gốc trước khi sửa các ô đó.
