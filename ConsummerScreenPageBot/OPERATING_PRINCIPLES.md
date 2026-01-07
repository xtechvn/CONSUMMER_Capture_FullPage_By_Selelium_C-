## Nguyên lý hoạt động của ConsummerScreenPageBot

### Kiến trúc tổng quan
- Ứng dụng console .NET (Selenium + ChromeDriver) chạy 1 tiến trình trình duyệt và lắng nghe RabbitMQ để nhận việc.
- Cấu hình chính lấy từ `App.config`: thông tin Rabbit (`RabbitHost`, `RabbitQueue`, `RabbitQueueImageAnalyze`, `RabbitQueueScreenLink`, SSL, vhost, user/pass), chế độ headless (`is_headless`), log path (`PATHLOG`), danh sách website, chế độ publish raw ảnh (`AnalyzePublishRaw`), nhiệm vụ bot (`task`).
- Kết quả chụp và cắt ảnh được lưu ở `screenshots/<hostname>/<device>/` và đồng thời push lên queue phân tích ảnh.

### Luồng khởi động
1) Thiết lập UTF-8 cho console.  
2) Tự tìm đường dẫn Chrome (macOS/Linux/Windows) và cấu hình `ChromeOptions`: user-agent, tắt automation flag, tạo profile riêng, chống crash headless (`--no-sandbox`, `--disable-dev-shm-usage`, v.v.).  
3) Khởi tạo `ChromeDriver` với log chi tiết.  
4) Mở kết nối RabbitMQ tới queue đầu vào `RabbitQueue` và đăng ký consumer (prefetch 1).  
5) Tiến trình giữ sống vĩnh viễn để nhận job liên tục.

### Định dạng job từ queue chính
- Payload JSON chứa tối thiểu:  
  - `link_web` (khi `task=screen_banner`) hoặc `link_click_banner` (khi `task` khác)  
  - `slice` (số phần cắt ảnh, mặc định 5/10)  
  - `quanlity_image` (chất lượng JPEG, mặc định 70)  
  - `device` (`"1"` desktop, `"2"` mobile, có thể `"1,2"`)  
  - `retry_screen_page` (số lần chụp lại)  
- Toàn bộ payload được lưu tạm vào `lastJobParams` để tái sử dụng khi build message phân tích.

### Xử lý mỗi message
1) Chọn device:
   - Desktop: viewport ~1920x1080 (`SetupDesktopDevice`).
   - Mobile: emulate 375x667 + UA mobile (`SetupMobileDevice`).
2) Điều hướng URL bằng `TryNavigate` (GoToUrl + fallback JS redirect, đợi ready state).
3) Gọi bộ xử lý:
   - Desktop: `ScreenDesktop.ProcessWebsite`
   - Mobile : `ScreenMobile.ProcessWebsite`
4) Ack message sau khi xử lý xong.

### Chụp trang chính (Desktop/Mobile)
`ProcessWebsite`:
- Xóa ảnh cũ của host.  
- Scroll kích hoạt lazy-load, thăm dò banner.  
- `CaptureSegmentScreenshots`:
  - Dùng CDP `Page.captureScreenshot` full page (fallback `ITakesScreenshot`), giới hạn chiều cao 30.000px.
  - Cắt ảnh thành N segment, lưu JPEG với `AdCapture.SaveImageCompressed`.
  - Publish mỗi segment lên queue phân tích qua `Program.TryPublishAnalyze(imageBytes, width, height)`.
- Sau khi cắt: gọi `ProcessIframesAndPushToQueue` (trừ khi `task=screen_link_page`).

### Quét banner/iframe và trang đích
`ProcessIframesAndPushToQueue`:
- Thu thập phần tử nghi quảng cáo: iframe + tập selector phổ biến (`GetCommonAdSelectors`) và đặc thù site.
- Lọc kích thước ≥120x30, tránh trùng bằng fingerprint.
- Với từng banner:
  - Cố gắng lấy link trực tiếp (href/data-href/data-url); nếu là iframe thì thử switch, click center hoặc open bằng JS.
  - Mở tab mới hoặc điều hướng trong tab; kiểm tra:
    - Link external (khác domain gốc) và trang có ảnh (`CheckPageHasImages`).
  - Chụp landing page: `CaptureFullPageScreenshotAsBase64`
    - Ghép 3 ảnh: đầu trang, cuối trang (scroll), và trang `/lien-he` nếu có menu/contact.
  - Push payload lên queue `RabbitQueueScreenLink` qua `PushIframeToQueue` (merge thêm `lastJobParams`, kèm `screenshot_base64`).
- Dọn tab thừa, quay về tab gốc sau mỗi banner.

### Publish ảnh cho AI phân tích
- `TryPublishAnalyze` đảm bảo kết nối publisher (`EnsureAnalyzePublisherReady`) tới `RabbitQueueImageAnalyze`.
- Payload xây bởi `Models/AnalyzePayloadBuilder.BuildAnalyzeBody`:
  - Nếu `AnalyzePublishRaw=1`: gửi raw bytes.
  - Ngược lại: JSON = `lastJobParams` + `screenshot_base64` + `width` + `height`.
- Message đặt `persistent=true`.

### Công cụ hỗ trợ
- `AdCapture`: cắt/lưu JPEG bằng ImageSharp, chụp element/iframe riêng lẻ và tự publish analyze.  
- `CaptureGenericBanners`/`HandleThanhNien`: selector đặc thù từng site.  
- `ScrollToBottomAndEnsureLazyContent`, `WaitForAdsLoaded`, `MaskSmallImages`: tối ưu nạp nội dung và dung lượng ảnh.  
- `ErrorWriter`: ghi log text theo ngày.  
- `TelegramService`: gửi cảnh báo lỗi qua Telegram.

### Thư mục và log
- Ảnh: `screenshots/<host>/<desktop|mobile>/...`  
- Log: `logs/error_yyyyMMdd.log`, `logs/chromedriver.log`  
- Cấu hình runtime: `App.config`

### Tóm tắt
Bot nhận job từ RabbitMQ, điều hướng URL trên Chrome (desktop/mobile), chụp ảnh full page thành nhiều segment và gửi lên queue phân tích. Sau đó bot dò banner/iframe, click để lấy trang đích, chụp thêm ảnh ghép và đẩy vào queue riêng phục vụ xử lý banner. Hệ thống tự khôi phục kết nối RabbitMQ, ghi log file và báo lỗi qua Telegram.

