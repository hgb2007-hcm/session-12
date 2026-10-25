# Chuẩn bị hồ sơ dự án trước khi triển khai hệ thống

## ***Phần 1. Phân tích: Mô tả sơ đồ cây cấu trúc Heading cho file SRS***

Để quản lý tài liệu SRS tính năng "Lớp học ảo" dài và có đủ 10 tiêu đề cấp 2 (Heading 2), sơ đồ cây cấu trúc phân cấp Heading trong Word (hiển thị qua Navigation Pane) được tổ chức như sau:
  
 1. Giới thiệu (Heading 1)
 
 - 1.1 Mục đích (Heading 2)   
 - 1.2 Phạm vi (Heading 2)
 - 1.3 Định nghĩa, viết tắt (Heading 2)   
 - 1.4 Tài liệu tham khảo (Heading 2)   

 2. Mô tả tổng quan (Heading 1) 
 - 2.1 Tổng quan hệ thống (Heading 2)
 - 2.2 Vai trò người dùng (Heading 2)

 3. Mô tả chức năng lớp học ảo (Heading 1)
 - 3.1 Tạo và quản lý lớp học (Heading 2)   
 - 3.2 Tham gia lớp học (Heading 2)   
 - 3.3 Tính năng trong lớp học (Heading 2)

4. Yêu cầu phi chức năng (Heading 1)
5. Ràng buộc & Giả định (Heading 1)
6. Phụ lục (Heading 1)

## ***Phần 2. Triển khai: Bộ 3 tài liệu (Word, Excel, PowerPoint)***

**1. Tài liệu Word: Đặc tả yêu cầu phần mềm (SRS) – Tính năng "Lớp học ảo"**

Tên file gợi ý: PTIT_Learning_SRS_LopHocAo.docx

- Cách thực hiện:
  - Sử dụng các mức Heading 1 và Heading 2 đúng chuẩn để Word tự động nhận diện.
  - Mở thẻ View > Navigation Pane để kiểm soát toàn bộ 10 tiêu đề cấp 2.
  - Trước khi lưu, thực hiện cập nhật lại mục lục tự động bằng cách nhấp chuột phải vào bảng mục lục và chọn Update Field > Update entire table (Cập nhật toàn bộ bảng) để đồng bộ số trang.

  **2. Bảng Excel: Quản lý danh sách học viên đăng ký đa nguồn**

  Tên file gợi ý: PTIT_Learning_QuanLyHocVien.xlsx

   
- Quy tắc chuẩn hóa dữ liệu:

  - Hàng đầu tiên là Header Row duy nhất chứa các tiêu đề cột (STT, Họ và tên, Email, SĐT, Nguồn đăng ký, Khóa học quan tâm, Ngày đăng ký, Trạng thái). Tuyệt đối không chèn hàng tổng cộng (Total) ở giữa vùng dữ liệu.   


- Cách dùng AutoFilter (Ctrl + Shift + L) để lọc học viên theo yêu cầu:
  - Nhấn tổ hợp phím Ctrl + Shift + L tại vùng dữ liệu để bật bộ lọc AutoFilter.
  - Tại cột Khoa/Bộ môn (hoặc lọc theo điều kiện điểm chuyên cần dưới 7 và khoa CNTT), bấm vào nút mũi tên lọc của cột tương ứng.
  - Chọn các tiêu chí lọc (ví dụ: Khoa: CNTT và Điểm chuyên cần < 7) để hiển thị chính xác danh sách theo yêu cầu đề bài.

**3. Slide PowerPoint: Tổng quan nền tảng PTIT-Learning**

Tên file gợi ý: PTIT_Learning_TongQuan.pptx

- Thiết kế và phối màu:
 
  - Sử dụng Combo màu tối: Nền màu xanh Navy đậm, chữ màu trắng sáng, và các điểm nhấn (Accent) dùng màu cam nổi bật để đảm bảo Contrast Ratio (tỷ lệ tương phản) cao, giúp người xem dễ quan sát trong phòng họp hoặc hội trường.

  - Bố cục chia slide thành chuẩn các vùng nội dung rõ ràng (Tiêu đề, khối lợi thế nổi bật, biểu đồ tăng trưởng học viên và cơ cấu nguồn đăng ký dạng biểu đồ tròn/donut).      

## ***Phần 3. Kiểm thử: Xử lý lỗi không tìm thấy tên học viên do khoảng trắng thừa bằng hàm TRIM***

Khi thực hiện lọc dữ liệu hoặc tìm kiếm một sinh viên bằng tên nhưng bảng Excel không trả về kết quả (mặc dù bạn nhìn thấy rõ tên đó xuất hiện trên danh sách), nguyên nhân lớn nhất là do dữ liệu bị dính khoảng trắng thừa (dư dấu cách ở đầu, ở cuối hoặc giữa các từ).

 **Quy trình xử lý bằng hàm TRIM:**

 - Tạo cột phụ: Chèn thêm một cột phụ bên cạnh cột tên gốc (ví dụ: cột Tên chuẩn hóa).

 - Áp dụng hàm TRIM: Nhập công thức tại ô tương ứng để loại bỏ khoảng trắng thừa:

 - *Excel:* =TRIM(B2)
 - Kéo công thức: Kéo công thức xuống cho toàn bộ danh sách học viên.
 - Thay thế dữ liệu cũ: Copy cột giá trị mới sau khi đã dùng TRIM, sau đó dán đè (Paste as Values) ngược lại vào cột tên gốc để làm sạch hoàn toàn dữ liệu, giúp việc lọc hoặc tra cứu hoạt động bình thường trở lại.