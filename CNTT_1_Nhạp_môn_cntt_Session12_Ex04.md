# Chuẩn hóa quy trình quản lý bệnh án điện tử 

## ***Phần 1. Phân tích: Phân biệt Yêu cầu chức năng và Phi chức năng***

Trong phát triển hệ thống phần mềm y tế, hai khái niệm này được phân biệt rõ ràng như sau:  

 - Yêu cầu chức năng (Functional Requirements - Khám bệnh): Là những tính năng nghiệp vụ trực tiếp phục vụ quy trình khám chữa bệnh của bác sĩ và bệnh nhân.
  
   - Ví dụ: Tính năng đăng ký lịch khám trực tuyến, nhập kết quả chẩn đoán, kê đơn thuốc điện tử, và tạo/cập nhật hồ sơ bệnh án điện tử.
- Yêu cầu phi chức năng (Non-functional Requirements - Bảo mật dữ liệu): Là các tiêu chuẩn về chất lượng, độ an toàn, hiệu suất và khả năng bảo vệ hệ thống (không trực tiếp thực hiện nghiệp vụ khám nhưng quyết định sự sống còn của hệ thống).

   - Ví dụ: Mã hóa dữ liệu bệnh nhân đạt chuẩn y tế, phân quyền nghiêm ngặt (chỉ bác sĩ điều trị mới được xem bệnh án chi tiết), sao lưu dữ liệu định kỳ trên cloud, và tốc độ tải trang dưới 2 giây

## ***Phần 2. Triển khai: Hệ thống Clinic-Cloud tích hợp***

**1. Tài liệu Word: SRS Module "Bệnh án điện tử" & Bảo mật**

 Tên file gợi ý: StartupClinic_SRS_BenhAnDienTu.docx


**Cấu trúc nội dung chính:**

1. Giới thiệu tổng quan: Mục tiêu hiện đại hóa lưu trữ hồ sơ y tế cho phòng khám.
2. Yêu cầu chức năng: Quy trình tiếp nhận bệnh nhân, khám lâm sàng và lưu trữ bệnh án.
3. Yêu cầu phi chức năng (Bảo mật): Các tiêu chuẩn phân quyền, mã hóa dữ liệu bệnh nhân và nhật ký truy cập (Audit Log).
4. Phụ lục & Thuật ngữ y khoa.
----

**2. Bảng Excel: Quản lý 1000 hồ sơ bệnh nhân (Database Mindset)** 

Tên file gợi ý: StartupClinic_QuanLyBenhNhan.xlsx

Cách xử lý bẫy ô trống (Empty Cells) khi dùng AutoFilter:

 - Vấn đề: Nếu bảng dữ liệu có hàng nghìn ô trống do bệnh nhân không điền đủ thông tin, khi dùng tính năng lọc dữ liệu (AutoFilter), Excel có thể hiểu nhầm vùng dữ liệu bị ngắt quãng và chỉ lọc đến ô trống đầu tiên.

 - Cách khắc phục:

   - Tuyệt đối không để dòng trống hoặc cột trống cắt ngang giữa bảng dữ liệu.  

   - Đối với các ô thông tin bị thiếu, nên điền giá trị mặc định (ví dụ: "Chưa cập nhật" hoặc "N/A") thay vì để trống hoàn toàn.

   - Trước khi bấm bật AutoFilter, hãy bôi đen (select) toàn bộ vùng bảng dữ liệu thay vì chỉ đứng ở một ô bất kỳ để Excel nhận diện chính xác toàn bộ 1000 dòng.
   -----

**3. Slide PowerPoint: Báo cáo số liệu dịch bệnh & Tổng quan hệ thống***  

Tên file gợi ý: StartupClinic_BaoCaoTongQuan.pptx

Thiết kế:

- Sử dụng Combo màu tối (Nền xanh Navy, chữ trắng sáng, điểm nhấn màu xanh dương/cam đặc trưng của y tế công nghệ) để đảm bảo độ tương phản cao.

- Tiêu đề slide viết dưới dạng kết luận (Assertion Title): "Hệ thống Clinic-Cloud giúp tối ưu 100% thời gian tra cứu hồ sơ và bảo mật tuyệt đối dữ liệu bệnh nhân".

- Trực quan hóa: Sử dụng Donut Chart (Biểu đồ vành khuyên tối giản) để thể hiện tỷ lệ các nhóm bệnh hoặc cơ cấu nguồn bệnh nhân đến khám, giúp slide gọn gàng, hiện đại và không bị rối mắt.
---

## ***Phần 3. Kiểm thử: Khắc phục lỗi gán nhãn Heading Styles nhầm bằng Ctrl + Shift + N***

Khi soạn thảo tài liệu SRS, nếu bạn vô tình bôi đen một đoạn văn bản dài (thậm chí cả một đoạn mô tả bệnh án) rồi bấm nhầm thành Heading 1 hoặc Heading 2, đoạn văn đó sẽ bị đưa vào khung Navigation Pane và làm hỏng cấu trúc Mục lục tự động.

**Cách khắc phục nhanh chóng bằng phím tắt:**
1. Bôi đen (Select) chính xác đoạn văn bản đang bị gán nhãn Heading nhầm.
2. Nhấn tổ hợp phím Ctrl + Shift + N trên bàn phím.
3. Kết quả: Thao tác này là phím tắt chuyển nhanh về định dạng Normal Style mặc định của Word, ngay lập tức đưa đoạn văn bản đó trở lại thành đoạn văn thường, loại bỏ nó khỏi bảng mục lục mà không cần phải bấm tìm kiếm trong danh sách Style phức tạp.