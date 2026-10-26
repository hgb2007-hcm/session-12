#  Hoàn thiện bộ hồ sơ triển khai nền tảng Smart-Home

## ***Phần 1. Phân tích: So sánh đặc điểm bảng Excel "Xấu" và bảng "Đúng" chuẩn Database***

 Cấu trúc ô : 
 
 - Có chứa các ô bị gộp (Merge Cell) ở phần tiêu đề hoặc giữa bảng, làm phá vỡ cấu trúc lưới.
 - Tuyệt đối không gộp ô trong vùng dữ liệu chính, mỗi ô là một giá trị độc lập.

Dòng/Cột trống :

- Chứa các hàng trống hoặc cột trống xen kẽ giữa các bản ghi để ngắt quãng trực quan.

- Liền mạch tuyệt đối, không có dòng trống hay cột trống ngắt quãng giữa các dữ liệu.

Tiêu đề bảng :

- Có nhiều hàng tiêu đề phụ phức tạp, chèn các hàng tổng cộng (Total) nằm ngay giữa dữ liệu.

- Chỉ có duy nhất một hàng Header Row ở trên cùng chứa tên các trường dữ liệu (Field).  

Khả năng xử lý :

- Lệnh lọc dữ liệu (AutoFilter) hoặc hàm tính toán (VLOOKUP) thường bị lỗi, bị ngắt quãng hoặc không quét hết bảng.

- Hoạt động trơn tru với mọi công cụ tự động hóa, bộ lọc, Pivot Table và hàm Excel.

## ***Phần 2. Triển khai: Toàn bộ học liệu dự án Smart-Home***

**1. Tài liệu Word: Hồ sơ dự án Smart-Home (Cấu trúc SRS)**

Tên file gợi ý: SmartHome_HoSoDuAn.docx

Cấu trúc phân cấp Heading (từ Heading 1 đến Heading 3):

 -  Giới thiệu chung (Heading 1)

   -  Mục tiêu dự án Smart-Home (Heading 2)

   -  Giải pháp nhà thông minh cho cuộc sống hiện đại (Heading 3)

   -  Định hướng phát triển thị trường (Heading 3)

** Nền tảng đấu giá nhà đất trực tuyến (Heading 1)** 

- Tính năng quản lý danh mục bất động sản (Heading 2)

-  Quy trình đấu giá và giao dịch an toàn (Heading 2)    

**Yêu cầu hệ thống và Bảo mật (Heading 1)** 

**2. Bảng Excel: Quản lý danh sách nhà đất & Tra cứu hoa hồng môi giới**

- Tên file gợi ý: SmartHome_QuanLyNhaDat.xlsx

- Cấu trúc dữ liệu: Gồm các trường (Field) rõ ràng: STT, Tên BĐS, Loại BĐS, Giá bán (VNĐ), Mã MG, Môi giới, Tỷ lệ hoa hồng, Hoa hồng (VNĐ).   

- Công thức VLOOKUP tra cứu tỷ lệ hoa hồng từ bảng phụ:
Tại ô tính tỷ lệ hoa hồng tương ứng của dòng đầu tiên, ta sử dụng hàm VLOOKUP kết hợp cố định vùng tham chiếu bằng dấu $: **=VLOOKUP(E2, $H$4:$I$7, 2, FALSE)**

**3. Slide PowerPoint: Pitch Deck thuyết phục nhà đầu tư (Rule 10/20/30)**

- Tên file gợi ý: SmartHome_PitchDeck.pptx

- Quy tắc trình bày:

  - Áp dụng Rule 10/20/30: Tối đa không quá 10 slide, thời gian thuyết trình trong vòng 20 phút, kích thước chữ tối ưu.   

  - Mỗi Slide - Một Thông điệp (Assertion Title): Tiêu đề slide thể hiện trực diện kết luận thay vì chỉ đặt tên chủ đề thô.

  - White Space (Khoảng trắng): Bố cục mỗi slide có khoảng trống thoáng đãng, kết hợp biểu đồ tăng trưởng và hình ảnh minh họa trực quan để làm nổi bật các ý chính.   

## ***Phần 3. Kiểm thử: Khắc phục lỗi trôi vùng dữ liệu khi kéo công thức bằng dấu $ (Dollar Sign)***

Khi thực hiện kéo công thức VLOOKUP xuống các dòng phía dưới cho toàn bộ danh sách, nếu ta không cố định vùng bảng tra cứu bằng dấu $, vùng tham chiếu sẽ bị dịch chuyển (trôi) lệch dần xuống dưới theo từng dòng (ví dụ từ $H$4:$I$7 bị trôi thành $H$5:$I$8, $H$6:$I$9...). Hậu quả là các dòng phía dưới sẽ bị lỗi #N/A do không tìm thấy dữ liệu trong phạm vi đã bị lệch.

Cách khắc phục bằng dấu $:

- Thêm dấu đô la $ vào trước tên cột và số hàng của vùng tham chiếu trong công thức (ví dụ: đổi từ H4:I7 thành $H$4:$I$7).

- Ý nghĩa: Dấu $ hoạt động như một chiếc "khóa chết", giúp cố định toàn bộ tọa độ ô bắt đầu và kết thúc của bảng tra cứu. Dù bạn có kéo công thức xuống hàng nghìn dòng bên dưới, vùng bảng phụ tra cứu vẫn đứng yên tuyệt đối, đảm bảo hàm trả về kết quả chính xác 100%.