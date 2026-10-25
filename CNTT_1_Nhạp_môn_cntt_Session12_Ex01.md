 # Chuẩn hóa quy trình đối soát giao dịch ví điện tử

 ## ***Phần 1. Phân tích: Xác định Record và Field trong bảng Excel giao dịch***

 - Trong bảng dữ liệu Excel đối soát giao dịch của ví điện tử PayKX, các khái niệm được hiểu đơn giản như sau:
   - Field (Cột / Trường dữ liệu): Là các thuộc tính mô tả thông tin của giao dịch. Ví dụ như: Mã giao dịch (Transaction ID), Mã người dùng (User ID), Tên người dùng (User Name), Số tiền giao dịch (Amount), Phí đối tác (Bank Fee), và Trạng thái (Status).

   - Record (Dòng / Bản ghi): Là toàn bộ thông tin chi tiết của một giao dịch cụ thể nằm trên một hàng ngang. Ví dụ, dòng thông tin của bạn Nguyễn Văn A nạp 50.000đ vào lúc 08:00 ngày 25/10/2026 chính là một Record hoàn chỉnh.

## ***Phần 2. Triển khai: Bộ 3 tài liệu (Word, Excel, PowerPoint)***

 **1. Tài liệu Word: Cấu trúc SRS (Tài liệu đặc tả yêu cầu phần mềm) cho tính năng Hoàn tiền ví điện tử PayKX**

 - Tên file gợi ý: PayKX_SRS_HoanTien.docx

 - Cấu trúc phân cấp Heading (dự kiến dài 5 trang):
   - ***Trang 1: Trang bìa (Tên dự án: PayKX, Trưởng dự án: Sinh viên thực hiện, Ngày tháng).***

   - ***Trang 2: Mục lục tự động (Auto TOC - sử dụng tính năng References > Table of Contents của Word).***

   - ***Trang 3 (Heading 1: 1. Giới thiệu chung):***
      - Heading 2: 1.1 Mục đích xây dựng tính năng hoàn tiền cho sinh viên.
      - Heading 2: 1.2 Phạm vi áp dụng trên ứng dụng PayKX. 
   - ***Trang 4 (Heading 1: 2. Mô tả chi tiết yêu cầu chức năng):***
      - Heading 2: 2.1 Luồng hoạt động của tính năng hoàn tiền (Cashback).
      - Heading 2: 2.2 Quy tắc tính toán tỷ lệ hoàn tiền theo từng dịch vụ (Nạp tiền điện thoại, thanh toán ký túc xá, mua vé xe buýt).
   - ***Trang 5 (Heading 1: 3. Giao diện và Xử lý ngoại lệ):***
      - Heading 2: 3.1 Thiết kế giao diện thông báo lịch sử nhận tiền hoàn.
      - Heading 2: 3.2 Xử lý lỗi khi giao dịch bị hủy nhưng tiền hoàn chưa về tài khoản.


**2. Bảng Excel: Đối soát giao dịch và tra cứu thông tin**

* Tên file gợi ý: PayKX_DoiSoat_Thang10.xlsx

* Quy tắc chuẩn hóa dữ liệu:
  - Tuyệt đối không dùng Merge Cell trong vùng bảng dữ liệu chính để tránh làm lệch hàm.
  - Không để dòng trống hay cột trống ngắt quãng giữa các bản ghi.
  - Dùng hàm TRIM để làm sạch dữ liệu tên hoặc mã người dùng bị dính khoảng trắng thừa (ví dụ: =TRIM(A2)).


***Công thức VLOOKUP tra cứu tên người dùng:
Giả sử bảng tra cứu mã người dùng và tên nằm ở sheet khác từ ô A$2:B$100, để lấy tên người dùng chính xác bằng mã giao dịch/mã người dùng ở cột B2, ta dùng công thức:***

- =VLOOKUP(TRIM(B2), Sheet2!$A$2:$B$100, 2, FALSE)

**3. Slide PowerPoint: Báo cáo doanh thu dòng tiền PayKX**

- Tên file gợi ý: PayKX_BaoCaoDoanhThu.pptx

- Áp dụng Quy tắc 10/20/30: Tối đa khoảng 10 slide, trình bày trong 20 phút, và kích thước font chữ tối thiểu là 30pt để người xem dễ đọc.

- Thiết kế Assertion Title (Tiêu đề khẳng định kết quả):
  - Slide 1: "Doanh thu ví điện tử PayKX tăng 15% nhờ ưu đãi nạp điện thoại học kỳ mới"
  - Slide 2: "Chi phí đối tác ngân hàng giảm nhờ tối ưu hóa cổng thanh toán tự động"
- Trực quan hóa: Thay vì chèn bảng số liệu thô nhiều chữ số, slide sử dụng Biểu đồ cột (Column Chart) thể hiện sự tăng trưởng dòng tiền qua từng tuần để ban giám khảo nhìn vào là hiểu ngay xu hướng.

## ***Phần 3. Kiểm thử: Quy trình 4 điểm xử lý lỗi #N/A khi tra cứu VLOOKUP***

Khi thực hiện hàm VLOOKUP mà Excel trả về lỗi #N/A (tức là không tìm thấy giá trị), mình sẽ kiểm tra lần lượt theo quy trình 4 bước sau:

- Kiểm tra khoảng trắng thừa (Leading/Trailing Spaces):
  
  - Nguyên nhân: Dữ liệu mã tra cứu có thể bị dính dấu cách ẩn ở đầu hoặc cuối (ví dụ "KH01 " khác với "KH01").
  - Cách xử lý: Dùng hàm TRIM() bao bọc giá trị dò tìm (như =VLOOKUP(TRIM(B2), ...)).

- Kiểm tra định dạng dữ liệu (Data Type Mismatch):
  - Nguyên nhân: Một bên ở bảng chính là dạng Text (chuỗi văn bản có dấu nháy đơn ' ở đầu), còn bên bảng tra cứu lại là dạng Number (số). Excel sẽ không khớp được.

  - Cách xử lý: Đồng bộ định dạng về chung một kiểu (chuyển tất cả sang Text bằng cách dùng hàm TEXT() hoặc sang Number).

- Kiểm tra phạm vi bảng tham chiếu (Table Array Range):

  - Nguyên nhân: Quên cố định vùng dữ liệu tham chiếu bằng dấu đô la (ví dụ quên viết thành $A$2:$B$100), dẫn đến khi kéo công thức xuống dưới, vùng tìm kiếm bị dịch chuyển lệch đi.
  - Cách xử lý: Nhấn phím F4 để khóa cố định vùng tham chiếu trong công thức.

 - Kiểm tra vị trí cột trả về (Col_index_num):

   - Nguyên nhân: Nhập số thứ tự cột cần lấy dữ liệu lớn hơn số cột thực tế có trong vùng tham chiếu (Table Array). 

   - Cách xử lý: Đếm lại số cột từ trái sang phải trong vùng chọn để điền đúng số thứ tự cột chứa thông tin cần lấy.
