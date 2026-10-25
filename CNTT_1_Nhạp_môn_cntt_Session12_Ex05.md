# Chuẩn hóa quy trình báo cáo sự cố an ninh mạng

## **Phần 1. Phân tích: Cơ chế "Google Search của bảng tính" thông qua tính năng AutoFilter**

Trong Excel, tính năng AutoFilter (Ctrl + Shift + L) thường được ví von như một chiếc "Google Search thu nhỏ" dành cho bảng dữ liệu vì cách vận hành cực kỳ thông minh của nó:

- Thanh tìm kiếm nhanh (Search Box): Khi bật bộ lọc, mỗi cột dữ liệu đều có một ô tìm kiếm. Ta có thể gõ trực tiếp từ khóa (ví dụ: gõ tên một loại mã độc như "Emotet" hoặc một địa chỉ IP đáng ngờ) để Excel lập tức quét toàn bộ hàng nghìn dòng và lọc ra đúng kết quả cần tìm, y hệt như cách ta gõ từ khóa lên thanh tìm kiếm Google.

- Lọc đa điều kiện (Multi-criteria Filtering): Cho phép kết hợp nhiều tiêu chí cùng lúc (ví dụ: vừa lọc các sự cố có mức độ là "Cao", vừa lọc đúng khung giờ tấn công ban đêm).

- Ẩn thông tin nhiễu: Thay vì phải đọc thủ công từng dòng một trong hàng nghìn bản ghi log bảo mật, AutoFilter giúp ẩn đi các dữ liệu không liên quan và chỉ hiển thị đúng "bản ghi vàng" mà chuyên viên an ninh đang cần phân tích.

## **Phần 2. Triển khai: Báo cáo NetGuard hoàn chỉnh**

**1. Tài liệu Word: Báo cáo sự cố an ninh mạng (Incident Report)** 

Tên file gợi ý: NetGuard_IncidentReport.docx

Cấu trúc & Định dạng:

- Sử dụng hệ thống phân cấp Heading Styles (Heading 1, Heading 2) để tổ chức rõ ràng các phần: Tổng quan sự cố, Phân tích log cảnh báo, Chi tiết mã độc, và Phương án khắc phục.

- Bật Navigation Pane (View > Navigation Pane) ở bên trái màn hình để dễ dàng kiểm soát, di chuyển giữa các chương của tài liệu dài hàng chục trang.

- Chèn Auto TOC (Mục lục tự động) ở trang đầu tiên và bấm Update Table trước khi lưu để đảm bảo số trang khớp chính xác.

**2. Bảng Excel: Tra cứu mã độc từ Cơ sở dữ liệu virus sang Nhật ký hệ thống**

- Tên file gợi ý: NetGuard_DoiSoatMaDoc.xlsx

- Công thức VLOOKUP kết hợp TRIM:
Để lấy thông tin tên mã độc từ bảng "Cơ sở dữ liệu virus" sang bảng "Nhật ký hệ thống" (tránh lỗi do dữ liệu thô bị dính khoảng trắng thừa ở mã định danh ID), ta sử dụng hàm TRIM bọc bên ngoài giá trị dò tìm:**=VLOOKUP(TRIM(D2), Database_Virus!$A$2:$D$100, 2, FALSE)**

## **3. Slide PowerPoint: Chiến lược phòng thủ 3 tầng**

 Tên file gợi ý: NetGuard_ChienLuocPhongThu.pptx

 Thiết kế & Phối màu:

 - Font chữ: Tuyệt đối không dùng Times New Roman, thay vào đó sử dụng các font chữ hiện đại, không chân như Montserrat hoặc Inter để hiển thị sắc nét, dễ đọc khi chiếu lên máy chiếu (Projector) trong phòng họp.

 - Bố cục & Nội dung: Sử dụng sơ đồ phân cấp (3 tầng phòng thủ: Protect - Detect - Respond) kết hợp quy tắc tối giản bullet points giúp người xem nắm bắt ngay lập tức chiến lược an ninh của trung tâm.   