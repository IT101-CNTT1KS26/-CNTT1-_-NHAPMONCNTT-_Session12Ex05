#  GIẢI BÀI TẬP: CHUẨN HÓA QUY TRÌNH BÁO CÁO SỰ CỐ AN NINH MẠNG (NETGUARD)

## PHẦN 1: PHÂN TÍCH — CƠ CHẾ "GOOGLE SEARCH CỦA BẢNG TÍNH" QUA TÍNH NĂNG AUTOFILTER

Trong quản trị an ninh mạng, khi hệ thống ghi nhận hàng nghìn dòng log cảnh báo sự cố từ tường lửa, máy chủ và các thiết bị endpoint, việc tìm kiếm thủ công là bất khả thi. Tính năng **AutoFilter** (`Ctrl + Shift + L`) hoạt động như một công cụ **"Google Search thu nhỏ"** ngay trong bảng tính nhờ các cơ chế sau:

1. **Cơ chế lập chỉ mục cột (Header Indexing):**
* Giống như công cụ tìm kiếm quét các trường dữ liệu trên web, AutoFilter tự động nhận diện hàng đầu tiên (Header Row) làm các trường thuộc tính (Field). Khi người dùng nhập từ khóa hoặc chọn điều kiện, hệ thống lập tức lọc qua cột tương ứng mà không làm xáo trộn trật tự bảng.


2. **Bộ lọc logic đa điều kiện (Logical Operators - AND/OR):**
* Cho phép kết hợp đồng thời nhiều tiêu chí khắt khe (ví dụ: *Lọc các sự cố có Mức độ = "Critical" VÀ Thời gian xảy ra trong 24 giờ gần nhất*). Cơ chế giao/hợp này giúp thu hẹp hàng nghìn bản ghi xuống chỉ còn vài dòng sự cố trọng điểm cần xử lý ngay lập tức.


3. **Cơ chế ẩn dòng bảo toàn dữ liệu (Row Hiding):**
* Khác với thao tác xóa (`Delete`) làm mất vĩnh viễn thông tin, AutoFilter chỉ tạm thời ẩn các dòng không thỏa mãn điều kiện, giúp bảo toàn toàn vẹn cấu trúc cơ sở dữ liệu gốc (Database Integrity) phục vụ cho việc lập báo cáo tổng hợp.



---

## PHẦN 2: TRIỂN KHAI — BÁO CÁO NETGUARD HOÀN CHỈNH

### 1. Tài liệu Word: Báo cáo sự cố an ninh mạng (Incident Report) & Auto TOC

* **Cấu trúc phân cấp Heading & Navigation Pane:** Tài liệu báo cáo gửi khách hàng được phân chia rõ ràng với `Heading 1` cho các chương lớn (Ví dụ: *1. Tổng quan sự cố an ninh mạng, 2. Phân tích chi tiết mã độc, 3. Khuyến nghị giải pháp phòng thủ*), kết hợp `Heading 2` và `Heading 3` cho các mục con. Cấu trúc này cho phép bật **Navigation Pane** (`Ctrl + F` chọn thẻ Headings) để điều hướng nhanh qua tài liệu hàng chục trang tức thì.
* **Mục lục tự động (Auto TOC):** Được chèn tại trang đầu thông qua tính năng `References > Table of Contents > Automatic Table 1`. Khi tài liệu phát sinh thêm nội dung làm thay đổi số trang, chỉ cần nhấp chuột phải vào bảng mục lục và chọn **Update Field > Update entire table** để đồng bộ chính xác tuyệt đối.

### 2. Bảng tính Excel: Tra cứu mã độc từ "Cơ sở dữ liệu virus" sang "Nhật ký hệ thống"

Để nối dữ liệu cảnh báo thô với tên định danh mã độc một cách chính xác, chống lỗi do khoảng trắng thừa hay mã phát sinh mới, ta sử dụng hàm `VLOOKUP` kết hợp `TRIM` và bẫy lỗi `IFERROR`:

* **Công thức chuẩn tại ô lấy tên mã độc (Ví dụ ô `D2`):**
```excel
=IFERROR(VLOOKUP(TRIM(B2), CoSoDuLieuVirus!$A$2:$D$500, 2, FALSE), "Mã độc chưa định danh")

```


* **Giải thích chi tiết công thức:**
* `TRIM(B2)`: Tự động loại bỏ các khoảng trắng thừa ẩn (ở đầu hoặc cuối chuỗi) trong mã độc thô tại ô B2, đảm bảo khớp hoàn hảo với danh mục.
* `CoSoDuLieuVirus!$A$2:$D$500`: Vùng dữ liệu tra cứu tại sheet cơ sở dữ liệu virus, được khóa tuyệt đối bằng ký hiệu đô la (`$`) để tránh bị trôi vùng khi kéo (fill down) công thức xuống hàng loạt.
* `2`: Lấy giá trị ở cột thứ 2 trong bảng tra cứu (tên định danh mã độc).
* `FALSE`: Yêu cầu tìm kiếm chính xác tuyệt đối (Exact Match).
* `IFERROR(..., "Mã độc chưa định danh")`: Xử lý mượt mà khi phát sinh mã độc mới lạ chưa có trong cơ sở dữ liệu, tránh hiển thị lỗi `#N/A` thô kệch trên báo cáo khách hàng.



### 3. Slide PowerPoint: Chiến lược phòng thủ 3 tầng (3-Tier Defense Strategy)

Bộ slide báo cáo chiến lược an ninh mạng được thiết kế tuân thủ nghiêm ngặt các nguyên tắc thiết kế hiện đại:

* **Quy định Font chữ:** **Tuyệt đối không dùng font Times New Roman**. Sử dụng các font chữ hiện đại, rõ nét khi chiếu trên Projector như **Montserrat** (cho tiêu đề) hoặc **Inter** (cho nội dung).
* **Tiêu đề kết luận (Assertion Title):**
* *Sai:* `Chiến Lược Phòng Thủ` hoặc `Báo Cáo An Ninh Mạng`
* *Đúng (Kết luận cốt lõi):* **`Mô Hình Phòng Thủ 3 Tầng Giúp NetGuard Ngăn Chặn 99.8% Các Cuộc Tấn Công Mạng Tự Động`**


* **Nội dung sơ đồ phân cấp Chiến lược phòng thủ 3 tầng (3-Tier Defense):**
1. **Tầng 1: Phòng thủ biên mạng (Perimeter Defense)**
* *Giải pháp:* Tường lửa thế hệ mới (NGFW), hệ thống ngăn chặn xâm nhập (IPS) và cổng lọc email tự động chống phishing.


2. **Tầng 2: Bảo vệ mạng nội bộ (Internal Network Defense)**
* *Giải pháp:* Phân đoạn mạng (Network Segmentation), giám sát lưu lượng ngang hàng (East-West Traffic) bằng SIEM/SOAR.


3. **Tầng 3: Bảo vệ điểm cuối và dữ liệu (Endpoint & Data Defense)**
* *Giải pháp:* Giải pháp EDR (Endpoint Detection and Response) trên từng máy trạm, mã hóa dữ liệu tĩnh và kiểm soát truy cập đặc quyền (PAM).