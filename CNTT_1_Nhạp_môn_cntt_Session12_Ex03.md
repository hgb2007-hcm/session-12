# Chuẩn hóa quy trình giao vận và tra cứu giá cước

## ***Phần 1. Phân tích: Mapping 4 tham số của VLOOKUP vào mô hình "Nhân viên kho tra hồ sơ"***

Để dễ hình dung cách hàm VLOOKUP hoạt động, ta có thể tưởng tượng công việc này giống hệt như một nhân viên kho đang tra cứu sổ sách hồ sơ:

- Tham số 1: lookup_value (Giá trị cần tìm) $\rightarrow$ Mã vận đơn trên bưu kiện: Là thông tin cụ thể mà nhân viên kho cầm trên tay (ví dụ: mã FM-1024) để bắt đầu đối chiếu.

- Tham số 2: table_array (Vùng dữ liệu tra cứu) $\rightarrow$ Cuốn sổ bảng giá cước chuẩn: Là toàn bộ danh mục / cuốn sổ lưu trữ thông tin cước phí đặt ở bàn làm việc (chứa các cột từ Mã vận đơn đến Giá cước).

- Tham số 3: col_index_num (Cột lấy dữ liệu) $\rightarrow$ Thứ tự cột "Giá cước" trong sổ: Là số thứ tự của cột chứa thông tin ta muốn tìm (tính từ trái sang phải trong cuốn sổ, ví dụ cột Giá cước nằm ở cột thứ 3).

- Tham số 4: range_lookup (Kiểu tìm kiếm FALSE / 0) $\rightarrow$ Yêu cầu khớp chính xác từng ký tự: Nhân viên kho phải so sánh chính xác 100% từng chữ cái và con số của mã vận đơn, không chấp nhận tra cứu gần đúng.

## ***Phần 2. Triển khai: Giải pháp công nghệ cho FastMove***

***1. Tài liệu Word: Quy trình vận hành kho chuẩn hóa***

 **Tên file gợi ý: FastMove_SOP_VanHanhKho.docx**

 **Thiết kế & Bố cục:**

 - Sử dụng bộ Heading Styles (Heading 1, Heading 2, Heading 3) đồng bộ theo màu sắc nhận diện thương hiệu của FastMove (ví dụ: Xanh da trời & Cam).

 - Áp dụng Auto TOC (Mục lục tự động) giúp quản lý tài liệu đồng bộ và dễ tra cứu.

 - Phân chia các mục rõ ràng: 1. Tiếp nhận hàng $\rightarrow$ 2. Phân loại & dán mã $\rightarrow$ 3. Xuất kho cho shipper.

 ***2. Bảng Excel: Tự động hóa tính giá cước (Áp dụng 5 Quy tắc Vàng)***

  - Tên file gợi ý: FastMove_BangTinhCuoc.xlsx

  - Tuân thủ 5 Quy tắc Vàng:

    1. Tuyệt đối không gộp ô (Merge Cell) ở bất kỳ vùng dữ liệu chính nào.
    2. Hàng đầu tiên chứa Header Row duy nhất.
    3. Không để dòng trống hay cột trống xen kẽ giữa các bản ghi.
    4. Mọi dữ liệu cùng một cột phải chuẩn hóa đúng định dạng (dạng Text cho Mã, dạng Number/Currency cho Giá tiền).
    5. Không chèn các hàng tổng cộng (Total) ở giữa bảng dữ liệu.

**Công thức VLOOKUP tra cứu cước phí từ Sheet2:**
Tại Sheet1 (danh sách 500 dòng đơn hàng), ở cột "Giá cước" (ô C2), nhập công thức có sử dụng dấu $ để cố định vùng bảng tra cứu ở Sheet2:

**Công thức Excel:** =VLOOKUP(A2, Sheet2!$A$2:$C$100, 3, FALSE)

***3. Slide PowerPoint: Hướng dẫn shipper nhận đơn & Báo cáo hiệu suất***

 - Tên file gợi ý: FastMove_HuongDanShipper.pptx
 - Áp dụng Miller's Law (Quy tắc trí nhớ ngắn hạn): 
   - Mỗi slide chỉ chứa tối đa 3 đến 4 ý (bullet points) giúp shipper nhanh chóng ghi nhớ thao tác nhận đơn mà không bị nạp quá tải thông tin.
 - Tiêu đề dạng Kết luận (Assertion Title):
   - Thay vì tiêu đề thô: "KPI giao hàng tháng 7"
   - Đặt tiêu đề kết luận: "Tỷ lệ giao hàng đúng giờ đạt 95% nhờ tự động hóa tra cứu giá cước"
 - Trực quan hóa: Sử dụng Bar Chart (Biểu đồ cột/thanh) báo cáo hiệu suất giao hàng, trong đó highlight (tô màu nổi bật) cột có thời gian giao hàng nhanh nhất để thu hút sự chú ý của người xem.

***Phần 3. Kiểm thử: Xử lý lỗi #N/A do ký tự lạ bằng hàm IFERROR***

Khi mã vận đơn bị dính các ký tự lạ hoặc chưa có trong bảng giá, hàm VLOOKUP sẽ trả về lỗi #N/A, làm mất thẩm mỹ và ảnh hưởng đến các ô tính tổng tiền phía dưới.

 **Cách xử lý kết hợp hàm IFERROR:** 

 Bao bọc công thức VLOOKUP bên trong hàm IFERROR để bẫy lỗi và hiển thị một thông báo rõ ràng: =IFERROR(VLOOKUP(A2, Sheet2!$A$2:$C$100, 3, FALSE), "Mã không hợp lệ")




