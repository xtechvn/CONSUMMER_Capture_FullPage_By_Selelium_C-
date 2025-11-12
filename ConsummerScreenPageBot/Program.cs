// Create By: cuonglv
// git add .
// git commit -m "update code mới"
// git push origin main
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Configuration;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using ConsummerScreenPageBot.Models;
using ConsummerScreenPageBot.Utils;
using ConsummerScreenPageBot.Device;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Jpeg;

namespace ConsummerScreenPageBot
{
    class Program
    {
        private static string startupPath = AppDomain.CurrentDomain.BaseDirectory.Replace(@"\bin\Debug\net8.0", @"\");

        // Queue name này chứa các link website cần screen
        public static string rabbit_queue_name = ConfigurationManager.AppSettings["RabbitQueue"] ?? "";
        // Queue name này chứa ảnh base64 để xử lý convert text
        public static string RabbitQueueAnalyze = ConfigurationManager.AppSettings["RabbitQueueImageAnalyze"] ?? "";
       
        // Queue name này dùng để nhận dữ liệu từ link click banner
        public static string RabbitQueueScreenLink = ConfigurationManager.AppSettings["RabbitQueueScreenLink"] ?? "";
        public static string rabbit_host = ConfigurationManager.AppSettings["RabbitHost"] ?? "";
        public static string rabbit_vhost = ConfigurationManager.AppSettings["RabbitVHost"] ?? "";
        public static int rabbit_port = Convert.ToInt32(ConfigurationManager.AppSettings["RabbitPort"] ?? "5672");
        public static string rabbit_username = ConfigurationManager.AppSettings["RabbitUserName"] ?? "";
        public static string rabbit_password = ConfigurationManager.AppSettings["RabbitPassword"] ?? "";    
        public static string rabbit_use_ssl = ConfigurationManager.AppSettings["RabbitUseSSL"] ?? "0";
        public static string LogPath = ConfigurationManager.AppSettings["PATHLOG"] ?? "logs";
        public static string is_headless = ConfigurationManager.AppSettings["is_headless"] ?? "0";
        public static string websites_config = ConfigurationManager.AppSettings["Websites"] ?? "";
        public static string analyze_publish_raw = ConfigurationManager.AppSettings["AnalyzePublishRaw"] ?? "0";

        // Publisher for analyze queue
        private static readonly object analyzePubLock = new object();
        private static IConnection? analyzeConnection;
        private static IModel? analyzeChannel;
        private static JObject? lastJobParams;

        /// <summary>
        /// Hàm chính khởi động ứng dụng
        /// 
        /// Cách thức hoạt động:
        /// 1. Thiết lập encoding UTF-8 cho console để hiển thị tiếng Việt đúng
        /// 2. Tìm đường dẫn Chrome binary tự động trên Windows/Linux/MacOS
        /// 3. Cấu hình ChromeOptions:
        ///    - Nếu headless mode: thiết lập các tham số cho server/docker (no-sandbox, disable-gpu, ...)
        ///    - Nếu không headless: mở Chrome full màn hình
        ///    - Tạo user profile riêng để tránh xung đột
        ///    - Thêm các tham số chống phát hiện automation
        /// 4. Khởi tạo ChromeDriver với logging chi tiết
        /// 5. Kết nối tới RabbitMQ queue để lắng nghe các job chụp ảnh website
        /// 6. Khi nhận được message từ queue:
        ///    - Parse JSON để lấy link_web, số lượng segment, chất lượng ảnh
        ///    - Điều hướng tới URL bằng TryNavigate
        ///    - Xử lý website bằng ProcessWebsite để chụp ảnh
        ///    - Gửi ảnh đã chụp lên queue analyze để AI xử lý
        /// 7. Giữ tiến trình luôn chạy để tiếp tục nhận job mới
        /// </summary>
        static void Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                var chrome_option = new ChromeOptions();

                // Fix 1: Add default path fallback for Chrome binary on missing/bad registry detection
                var chromeBinary = FindChromeBinaryPath();
                if (!string.IsNullOrWhiteSpace(chromeBinary) && File.Exists(chromeBinary))
                {
                    chrome_option.BinaryLocation = chromeBinary;
                    Console.WriteLine($"[Chrome] Using Chrome binary: {chromeBinary}");
                }
                else
                {
                    // Dòng bổ sung, cảnh báo khi Chrome không tìm thấy
                    Console.WriteLine("Warning: Chrome binary NOT found. Please check Chrome is installed or adjust the path.");
                    Console.WriteLine($"[Chrome] Attempting to use system default Chrome...");
                    // Không set BinaryLocation, để ChromeDriver tự tìm
                }

                // Fix 2: Cấu hình chống crash headless/windowless môi trường server/docker
                if (is_headless == "1")
                {
                    Console.WriteLine("[Chrome] Configuring headless mode...");
                    chrome_option.AddArgument("--headless=new");
                    chrome_option.AddArgument("--disable-gpu");
                    chrome_option.AddArgument("--window-size=1920,1080");
                    chrome_option.AddArgument("--disable-software-rasterizer");
                    chrome_option.AddArgument("--no-sandbox");
                    chrome_option.AddArgument("--disable-dev-shm-usage");
                    // Thêm options cho macOS
                    if (OperatingSystem.IsMacOS())
                    {
                        Console.WriteLine("[Chrome] Adding macOS-specific options...");
                        chrome_option.AddArgument("--disable-setuid-sandbox");
                        chrome_option.AddArgument("--disable-web-security");
                        chrome_option.AddArgument("--disable-features=VizDisplayCompositor");
                        chrome_option.AddArgument("--disable-background-timer-throttling");
                        chrome_option.AddArgument("--disable-backgrounding-occluded-windows");
                        chrome_option.AddArgument("--disable-renderer-backgrounding");
                        // Thêm option để tránh ProcessSingleton lock conflict
                        chrome_option.AddArgument("--remote-debugging-pipe");
                    }
                }
                else
                {
                    chrome_option.AddArgument("--start-maximized"); // set full man hinh  
                }

                // Use isolated user profile với unique ID để tránh lock conflict
                var userDataDir = Path.Combine(startupPath, "chrome-profile", Guid.NewGuid().ToString("N").Substring(0, 8));
                if (!Directory.Exists(userDataDir)) Directory.CreateDirectory(userDataDir);
                chrome_option.AddArgument($"--user-data-dir={userDataDir}");
                
                // Xóa lock file nếu tồn tại để tránh conflict
                try
                {
                    var lockFile = Path.Combine(userDataDir, "SingletonLock");
                    if (File.Exists(lockFile))
                    {
                        File.Delete(lockFile);
                    }
                }
                catch { }

                // Các option giả lập người dùng
                chrome_option.AddArgument("--disable-blink-features=AutomationControlled");
                chrome_option.AddExcludedArgument("enable-automation");
                chrome_option.AddAdditionalOption("useAutomationExtension", false);
                chrome_option.AddArgument("--disable-infobars");
                chrome_option.AddArgument("--no-first-run");
                chrome_option.AddArgument("--no-default-browser-check");
                chrome_option.AddArgument("--disable-background-networking");
                chrome_option.AddArgument("--disable-extensions");
                chrome_option.AddArgument("--disable-sync");
                chrome_option.AddArgument("--disable-component-update");
                chrome_option.AddArgument("--disable-client-side-phishing-detection");
                chrome_option.AddArgument("--disable-domain-reliability");
                chrome_option.AddArgument("--disable-renderer-backgrounding");
                chrome_option.AddUserProfilePreference("credentials_enable_service", false);
                chrome_option.AddUserProfilePreference("profile.password_manager_enabled", false);
                // Improve compatibility for newer Chrome versions and certain environments
                chrome_option.AddArgument("--remote-allow-origins=*");
                chrome_option.AddArgument("--disable-features=IsolateOrigins,site-per-process");
                chrome_option.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36");
                chrome_option.AcceptInsecureCertificates = true;
                chrome_option.PageLoadStrategy = PageLoadStrategy.Eager;

                // Fix 3: Check chrome driver compatibility or print its version
                var driverVersion = typeof(ChromeDriver).Assembly.GetName().Version;
                Console.WriteLine("ChromeDriver NuGet version: " + driverVersion);

             
                // Tạo ChromeDriverService và bật ghi log chi tiết để dễ debug khi lỗi phiên làm việc/khởi tạo
                var service = ChromeDriverService.CreateDefaultService();
                service.EnableVerboseLogging = true;
                service.LogPath = Path.Combine(LogPath, "chromedriver.log");

                // Fix 4: Chỉ cảnh báo DISPLAY trên hệ điều hành không phải Windows (Linux/Unix)
                try
                {
                    if (!OperatingSystem.IsWindows())
                    {
                        var display = Environment.GetEnvironmentVariable("DISPLAY");
                        if (string.IsNullOrWhiteSpace(display) && is_headless != "1")
                        {
                            Console.WriteLine("No DISPLAY found! Your environment probably needs to run with is_headless=1. Otherwise, Chrome cannot start without a display server.");
                        }
                    }
                }
                catch { }

                try
                {
                    Console.WriteLine("[Chrome] Starting ChromeDriver...");
                    using (var browers = new ChromeDriver(service, chrome_option, TimeSpan.FromMinutes(3)))
                    {
                        Console.WriteLine("[Chrome] ChromeDriver started successfully!");  
                         
                        #region WAITING QUEUE
                        var factory = new ConnectionFactory()
                        {
                            HostName = rabbit_host,
                            UserName = rabbit_username,
                            Password = rabbit_password,
                            VirtualHost = string.IsNullOrWhiteSpace(rabbit_vhost) ? "/" : rabbit_vhost,
                            Port = rabbit_port,
                            AutomaticRecoveryEnabled = true,
                            NetworkRecoveryInterval = TimeSpan.FromSeconds(5),    // Nếu mất kết nối RabbitMQ, thử kết nối lại sau mỗi 5 giây
                            //RequestedConnectionTimeout = TimeSpan.FromSeconds(30), // Thời gian timeout tối đa khi chờ thiết lập kết nối là 30 giây
                            //RequestedHeartbeat = TimeSpan.FromSeconds(30)          // Gửi tín hiệu heartbeat tới RabbitMQ mỗi 30 giây để kiểm tra kết nối còn sống
                        };
                        if (rabbit_use_ssl == "1" || rabbit_port == 5671)
                        {
                            factory.Ssl = new SslOption
                            {
                                Enabled = true,
                                ServerName = rabbit_host,
                                AcceptablePolicyErrors = System.Net.Security.SslPolicyErrors.None
                            };
                        }
                    using (var connection = factory.CreateConnection())
                    using (var channel = connection.CreateModel())
                    {
                        try
                        {
                            channel.QueueDeclare(queue: rabbit_queue_name,
                                                 durable: true,
                                                 exclusive: false,
                                                 autoDelete: false,
                                                 arguments: null);

                            channel.BasicQos(prefetchSize: 0, prefetchCount: 1, global: false);

                            Console.WriteLine(" [*] Waiting for messages on: " + rabbit_queue_name);

                            var consumer = new EventingBasicConsumer(channel);
                            consumer.Received += (sender, ea) =>
                            {
                                try
                                {
                                    var body = ea.Body.ToArray();
                                    var message = Encoding.UTF8.GetString(body);

                                    Console.WriteLine("Received banner data Queue: {0}", message + "----------");
                                    // Đọc link_web từ message (dạng JSON)
                                    string siteUrl = "";
                                    int segment_page = 10;
                                    var jobj = JObject.Parse(message);

                                    // Link web can chup
                                    siteUrl = jobj["link_web"]?.ToString() ?? "";

                                    // So luong segment chia trang de gui cho GEMINI phan tich hinh anh
                                    segment_page = jobj["slice"] != null ? jobj["slice"].ToObject<int>() : 5;

                                    // Chat luong anh xuat ra de OCR GEMINI phan tich hinh anh
                                    long jpegQuality = jobj["quanlity_image"] != null ? jobj["quanlity_image"].ToObject<long>() : 70;

                                    // Day la so lan chup. sau khi refresh trang se tinh la 1 lan chup
                                    int retry_screen_page = jobj["retry_screen_page"] != null ? jobj["retry_screen_page"].ToObject<int>() : 1;

                                    // Device: 1:PC, 2:Mobile (có thể là "1", "2", hoặc "1,2")
                                    string device = jobj["device"] != null ? jobj["device"].ToObject<string>() : "1";
                                    
                                    

                                    try { 
                                        lastJobParams = (JObject)jobj.DeepClone(); 
                                    } catch {
                                         lastJobParams = jobj; 
                                    }
                                   
                                    
                                    try
                                    {                                        
                                        // Parse device parameter: có thể là "1", "2", hoặc "1,2"
                                        var deviceList = device.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                        
                                        foreach (var dev in deviceList)
                                        {
                                            var deviceType = dev.Trim();
                                            try
                                            {
                                                Console.WriteLine($"Processing device type: {deviceType} for {siteUrl}");
                                                
                                                // Setup ChromeOptions cho device type
                                                IWebDriver driver = browers;
                                                bool needMobileSetup = deviceType == "2";
                                                
                                                if (needMobileSetup)
                                                {
                                                    // Setup mobile device emulation
                                                    SetupMobileDevice(browers);
                                                }
                                                else
                                                {
                                                    // Setup desktop (hoặc giữ nguyên nếu đã là desktop)
                                                    SetupDesktopDevice(browers);
                                                }
                                                
                                                Console.WriteLine("Navigate: " + siteUrl);
                                                if (!TryNavigate(browers, siteUrl, TimeSpan.FromSeconds(60), out var failReason))
                                                {
                                                    Console.WriteLine($"Navigation failed: {failReason}");
                                                    ErrorWriter.WriteLog(LogPath, "NavigateFail", $"{deviceType} - {siteUrl} => {failReason}");
                                                    continue;
                                                }
                                                
                                                // Gọi đúng class dựa trên device type
                                                if (deviceType == "2")
                                                {
                                                    // Mobile
                                                    ScreenMobile.ProcessWebsite(browers, siteUrl, segment_page, jpegQuality);
                                                }
                                                else
                                                {
                                                    // Desktop (mặc định)
                                                    ScreenDesktop.ProcessWebsite(browers, siteUrl, segment_page, jpegQuality);
                                                }
                                            }
                                            catch (Exception devEx)
                                            {
                                                Console.WriteLine($"Error processing device {deviceType} for {siteUrl}: {devEx.Message}");
                                                ErrorWriter.WriteLog(LogPath, "ProcessDevice", $"{deviceType} - {siteUrl} => {devEx}");
                                                TelegramService.PushLogToTelegram($"Error processing device {deviceType} for website: {siteUrl}", devEx);
                                            }
                                        }
                                    }
                                    catch (Exception siteEx)
                                    {
                                        Console.WriteLine($"Error processing {siteUrl}: {siteEx.Message}");
                                        ErrorWriter.WriteLog(LogPath, "ProcessWebsite", $"{siteUrl} => {siteEx}");
                                        TelegramService.PushLogToTelegram($"Error processing website: {siteUrl}", siteEx);
                                    }
                                    

                                    channel.BasicAck(deliveryTag: ea.DeliveryTag, multiple: false);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("error queue: " + ex.ToString());
                                    ErrorWriter.WriteLog(LogPath, "QueueError", ex.ToString());
                                    TelegramService.PushLogToTelegram("Queue Error occurred", ex);
                                }
                            };

                            channel.BasicConsume(queue: rabbit_queue_name, autoAck: false, consumer: consumer);

                            // Giữ tiến trình luôn lắng nghe queue, không tự động thoát
                            var hold = new ManualResetEventSlim(false);
                            hold.Wait();

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                            throw;
                        }
                    }
                    
                        #endregion
                        // Không tự động đóng trình duyệt để luôn sẵn sàng nhận job mới
                    }
                }
                catch (WebDriverException wdEx)
                {
                    Console.WriteLine("WebDriverException: " + wdEx.Message);  
                    ErrorWriter.WriteLog(LogPath, "SessionNotCreated", wdEx.ToString());
                    TelegramService.PushLogToTelegram("WebDriverException: Session Not Created", wdEx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.ToString());
                    ErrorWriter.WriteLog(LogPath, "Handle Error", ex.ToString());
                    TelegramService.PushLogToTelegram("Handle Error occurred", ex);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error (outer): " + ex.ToString());
                ErrorWriter.WriteLog(LogPath, "Handle Error(outer)", ex.ToString());
                TelegramService.PushLogToTelegram("Handle Error (outer) occurred", ex);
            }
        }

        /// <summary>
        /// Setup Chrome để giả lập Mobile device
        /// 
        /// Cách thức hoạt động:
        /// 1. Sử dụng Chrome DevTools Protocol để thiết lập mobile device emulation
        /// 2. Đặt viewport thành 375x667 (iPhone 12/13)
        /// 3. Đặt mobile = true, deviceScaleFactor = 2 (retina display)
        /// 4. Đặt User Agent Mobile (iPhone)
        /// </summary>
        private static void SetupMobileDevice(IWebDriver driver)
        {
            try
            {
                var chrome = driver as ChromeDriver;
                if (chrome != null)
                {
                    // Mobile device metrics
                    var metrics = new Dictionary<string, object>
                    {
                        { "mobile", true },
                        { "width", 375 },
                        { "height", 667 },
                        { "deviceScaleFactor", 2 },
                        { "scale", 1 }
                    };
                    chrome.ExecuteCdpCommand("Emulation.setDeviceMetricsOverride", metrics);

                    // Mobile User Agent (iPhone)
                    var userAgent = "Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Mobile/15E148 Safari/604.1";
                    chrome.ExecuteCdpCommand("Network.setUserAgentOverride", new Dictionary<string, object>
                    {
                        { "userAgent", userAgent }
                    });

                    Console.WriteLine("[Mobile] Device emulation setup completed");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting up mobile device: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "SetupMobileDevice", ex.ToString());
            }
        }

        /// <summary>
        /// Setup Chrome để giả lập Desktop device
        /// 
        /// Cách thức hoạt động:
        /// 1. Sử dụng Chrome DevTools Protocol để thiết lập desktop device
        /// 2. Đặt viewport thành 1920x1080
        /// 3. Đặt mobile = false, deviceScaleFactor = 1
        /// 4. Đặt User Agent Desktop (Windows)
        /// </summary>
        private static void SetupDesktopDevice(IWebDriver driver)
        {
            try
            {
                var chrome = driver as ChromeDriver;
                if (chrome != null)
                {
                    // Desktop device metrics
                    var metrics = new Dictionary<string, object>
                    {
                        { "mobile", false },
                        { "width", 1920 },
                        { "height", 1080 },
                        { "deviceScaleFactor", 1 },
                        { "scale", 1 }
                    };
                    chrome.ExecuteCdpCommand("Emulation.setDeviceMetricsOverride", metrics);

                    // Desktop User Agent
                    var userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36";
                    chrome.ExecuteCdpCommand("Network.setUserAgentOverride", new Dictionary<string, object>
                    {
                        { "userAgent", userAgent }
                    });

                    Console.WriteLine("[Desktop] Device emulation setup completed");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting up desktop device: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "SetupDesktopDevice", ex.ToString());
            }
        }

        /// <summary>
        /// Xử lý đặc biệt cho website VnExpress
        /// 
        /// Cách thức hoạt động:
        /// 1. Đợi trang web load hoàn toàn (readyState = "complete")
        /// 2. Đảm bảo các quảng cáo ở phần đầu trang (top ads) đã được render
        ///    - Cuộn lên đầu trang
        ///    - Lắc scroll nhỏ để kích hoạt observer
        ///    - Đếm số lượng quảng cáo hiển thị ở vùng top (từ 0 đến 900px)
        ///    - Đợi cho đến khi số lượng quảng cáo ổn định (không thay đổi trong 2 lần kiểm tra)
        /// 
        /// Hàm này tối ưu cho VnExpress vì trang này có nhiều quảng cáo lazy-load ở đầu trang
        /// </summary>
        private static void HandleVnExpress(IWebDriver driver)
        {
            var wait = new WebDriverWait(new SystemClock(), driver, TimeSpan.FromSeconds(15), TimeSpan.FromMilliseconds(250));
            try
            {
                wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").ToString() == "complete");
            }
            catch { }

            // Đảm bảo top ads render trước khi chụp
            EnsureTopAdsOnVnExpress(driver, TimeSpan.FromSeconds(8));            

        }

        /// <summary>
        /// Đảm bảo các quảng cáo ở phần đầu trang VnExpress đã được render hoàn toàn
        /// 
        /// Cách thức hoạt động:
        /// 1. Cuộn lên đầu trang (scrollTo(0,0))
        /// 2. Trong khoảng thời gian maxWait:
        ///    - Lắc scroll nhẹ (scroll 40px xuống rồi quay lại 0) để kích hoạt mutation observer
        ///    - Chạy JavaScript để đếm số lượng quảng cáo hiển thị:
        ///      * Tìm các selector: header banner, #banner_top, .gpt-ad, 1, ...
        ///      * Chỉ tính các quảng cáo có kích thước >= 120x30px và nằm trong vùng top (y < 900px)
        ///    - Nếu số lượng quảng cáo ổn định trong 2 lần kiểm tra liên tiếp => thoát
        ///    - Nếu không tìm thấy quảng cáo => đợi 220ms rồi kiểm tra lại
        ///    - Nếu số lượng thay đổi => reset counter và tiếp tục
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - maxWait: Thời gian tối đa để đợi quảng cáo load
        /// </summary>
        private static void EnsureTopAdsOnVnExpress(IWebDriver driver, TimeSpan maxWait)
        {
            var js = (IJavaScriptExecutor)driver;
            try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }

            var end = DateTime.UtcNow + maxWait;
            int stable = 0;
            int lastCount = -1;

            while (DateTime.UtcNow < end)
            {
                try
                {
                    // lắc scroll nhỏ để kích hoạt observer
                    try { js.ExecuteScript("window.scrollTo(0, 40);"); } catch { }
                    System.Threading.Thread.Sleep(120);
                    try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                }
                catch { }

                int visibleTopAds = 0;
                try
                {
                    var script = @"
                        const sels = [
                          'header .banner', 'header [id*=\\'banner\\']', 'header [class*=\\'banner\\']',
                          '#banner_top', '.banner-top', '.top-banner', '.leaderboard', '.top-ads', '.banner-leaderboard',
                          '[id^=\\'div-gpt-ad\\']', '[id*=\\'gpt-ad\\']', '.gpt-ad', '.dfp-ad', '.ad-slot', '.ad-container'
                        ].join(',');
                        const minW = 120, minH = 30, maxY = 900;
                        const y = window.scrollY || window.pageYOffset || 0;
                        let count = 0;
                        const list = Array.from(document.querySelectorAll(sels));
                        for (const el of list) {
                          try {
                            const r = el.getBoundingClientRect();
                            const w = Math.round(r.width), h = Math.round(r.height);
                            const top = Math.round(r.top + y);
                            if (w >= minW && h >= minH && top < maxY) count++;
                          } catch {}
                        }
                        const ifr = Array.from(document.querySelectorAll('iframe,frame'));
                        for (const el of ifr) {
                          try {
                            const r = el.getBoundingClientRect();
                            const w = Math.round(r.width), h = Math.round(r.height);
                            const top = Math.round(r.top + y);
                            if (w >= minW && h >= minH && top < maxY) count++;
                          } catch {}
                        }
                        return count;
                    ";
                    visibleTopAds = Convert.ToInt32(js.ExecuteScript(script) ?? 0);
                }
                catch { }

                if (visibleTopAds <= 0)
                {
                    System.Threading.Thread.Sleep(220);
                    continue;
                }

                if (visibleTopAds == lastCount) stable++; else { stable = 0; lastCount = visibleTopAds; }
                if (stable >= 2) break;

                System.Threading.Thread.Sleep(180);
            }
        }

        /// <summary>
        /// Xóa tất cả file ảnh cũ trong thư mục screenshots của host cụ thể
        /// 
        /// Cách thức hoạt động:
        /// 1. Tạo đường dẫn thư mục screenshots/<hostLabel> (ví dụ: screenshots/vnexpress.net)
        /// 2. Nếu thư mục không tồn tại => không làm gì
        /// 3. Lấy danh sách tất cả file trong thư mục
        /// 4. Duyệt qua từng file và xóa:
        ///    - Nếu xóa thành công => tiếp tục
        ///    - Nếu xóa lỗi => ghi log lỗi nhưng vẫn tiếp tục xóa file khác
        /// 
        /// Mục đích: Làm sạch thư mục trước khi chụp ảnh mới để tránh file cũ còn sót lại
        /// 
        /// Tham số:
        /// - hostLabel: Hostname của website (ví dụ: "vnexpress.net")
        /// </summary>
        private static void ClearHostScreenshots(string hostLabel)
        {
            try
            {
                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) return;

                var files = Directory.GetFiles(shotsDir);
                foreach (var f in files)
                {
                    try { File.Delete(f); }
                    catch (Exception ex)
                    {
                        ErrorWriter.WriteLog(LogPath, "ClearScreenshots.Delete", $"{hostLabel} => {f} => {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "ClearScreenshots", $"{hostLabel} => {ex}");
                TelegramService.PushLogToTelegram($"ClearScreenshots Error - {hostLabel}", ex);
            }
        }

       
        /// <summary>
        /// Chụp ảnh trang web bằng cách chia thành nhiều segment (phần) và chụp từng phần
        /// 
        /// Cách thức hoạt động:
        /// 1. Lấy kích thước thực tế của trang (width x height)
        ///    - Giới hạn chiều cao tối đa 30000px
        /// 2. Cuộn xuống cuối trang và đợi quảng cáo load hoàn toàn
        /// 3. Chụp ảnh toàn trang một lần bằng CDP:
        ///    - Đặt viewport theo kích thước tài liệu
        ///    - Chụp ảnh full page bằng Page.captureScreenshot
        ///    - Nếu CDP fail => fallback về ITakesScreenshot
        /// 4. Load ảnh full page vào ImageSharp để xử lý
        /// 5. Chia ảnh thành N segment bằng nhau (slice):
        ///    - Tính chiều cao mỗi segment: totalHeight / segmentCount
        ///    - Với mỗi segment:
        ///      * Cắt ảnh theo vị trí Y tương ứng (y = i * sliceHeight)
        ///      * Lưu segment với tên file: split{i+1}_YYYYMMDD_HHmmss_ffffff_guid.jpg
        ///      * Nén ảnh với chất lượng jpegQuality
        ///      * Gửi ảnh lên queue analyze để AI xử lý
        /// 6. Reset viewport về kích thước ban đầu
        /// 
        /// Ưu điểm: Chia nhỏ giúp AI dễ xử lý từng phần và giảm dung lượng mỗi ảnh
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - hostLabel: Hostname
        /// - segmentCount: Số lượng segment cần chia (mặc định 10)
        /// - jpegQuality: Chất lượng JPEG
        /// </summary>
        private static void CaptureSegmentScreenshots(IWebDriver driver, string hostLabel, int segmentCount, long jpegQuality)
        {
            try
            {
                try { Console.WriteLine($"[Segment] Start capture host={hostLabel}, segments={segmentCount}, quality={jpegQuality}"); } catch { }
                if (segmentCount <= 0) segmentCount = 3;

                var js = (IJavaScriptExecutor)driver;
                // Kích hoạt lazy-load và đợi quảng cáo hiển thị trước khi chụp full
                ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(8));
                WaitForAdsLoaded(driver, TimeSpan.FromSeconds(6));
                int pageWidth = 1920;
                int totalHeight = 3000;
                try
                {
                    pageWidth = Convert.ToInt32(js.ExecuteScript("return Math.max(document.documentElement.scrollWidth, document.body.scrollWidth, window.innerWidth || 0);"));
                    totalHeight = Convert.ToInt32(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight, window.innerHeight || 0);"));
                }
                catch { }
                if (totalHeight <= 0) totalHeight = 3000;
                totalHeight = Math.Min(totalHeight, 30000);

                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                // 1) Chụp full page một lần bằng CDP; fallback sang ITakesScreenshot nếu cần
                byte[] fullShotBytes = Array.Empty<byte>();
                var chrome = driver as ChromeDriver;
                bool fullOk = false;
                if (chrome != null)
                {
                    var metrics = new Dictionary<string, object>
                    {
                        { "mobile", false },
                        { "width", Math.Max(1, pageWidth) },
                        { "height", Math.Max(1, totalHeight) },
                        { "deviceScaleFactor", 1 },
                        { "scale", 1 }
                    };
                    try { chrome.ExecuteCdpCommand("Emulation.setDeviceMetricsOverride", metrics); } catch { }

                    try { chrome.ExecuteCdpCommand("Page.enable", new Dictionary<string, object>()); } catch { }

                    try
                    {
                        var args = new Dictionary<string, object>
                        {
                            { "format", "jpeg" },
                            { "quality", 100 },
                            { "captureBeyondViewport", true }
                        };
                        var result = chrome.ExecuteCdpCommand("Page.captureScreenshot", args) as IDictionary<string, object>;
                        if (result != null && result.TryGetValue("data", out var dataObj) && dataObj is string base64)
                        {
                            fullShotBytes = Convert.FromBase64String(base64);
                            fullOk = fullShotBytes != null && fullShotBytes.Length > 0;
                        }
                    }
                    catch { fullOk = false; }

                    if (!fullOk)
                    {
                        try
                        {
                            var shot = ((ITakesScreenshot)driver).GetScreenshot();
                            fullShotBytes = shot.AsByteArray;
                            fullOk = fullShotBytes != null && fullShotBytes.Length > 0;
                        }
                        catch { fullOk = false; }
                    }

                    try { chrome.ExecuteCdpCommand("Emulation.clearDeviceMetricsOverride", new Dictionary<string, object>()); } catch { }
                }
                else
                {
                    // No CDP: best effort viewport screenshot
                    try
                    {
                        var shot = ((ITakesScreenshot)driver).GetScreenshot();
                        fullShotBytes = shot.AsByteArray;
                        fullOk = fullShotBytes != null && fullShotBytes.Length > 0;
                    }
                    catch { fullOk = false; }
                }
                if (!fullOk || fullShotBytes == null || fullShotBytes.Length == 0)
                {
                    throw new Exception("Failed to capture full page screenshot for slicing.");
                }

                // 2) Cắt ảnh theo N phần từ fullShotBytes, luôn đảm bảo phần cuối chứa phần còn lại (footer)
                using (var ms = new MemoryStream(fullShotBytes))
                using (var fullImg = Image.Load<Rgba32>(ms))
                {
                    int sliceHeight = (int)Math.Ceiling(fullImg.Height / (double)segmentCount);
                    for (int i = 0; i < segmentCount; i++)
                    {
                        int y = i * sliceHeight;
                        int currentHeight = Math.Min(sliceHeight, Math.Max(1, fullImg.Height - y));
                        if (currentHeight <= 0) break;

                        using (var seg = fullImg.Clone(ctx => ctx.Crop(new SixLabors.ImageSharp.Rectangle(0, y, fullImg.Width, currentHeight))))
                        {
                            var fileName = $"split{i+1}_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.jpg";
                            var savePath = Path.Combine(shotsDir, fileName);
                            // Nén như cũ (scale 50%)
                            AdCapture.SaveImageCompressed(seg, savePath, 1.0, jpegQuality);

                            try
                            {
                                var compressedBytes = File.ReadAllBytes(savePath);
                                try { Console.WriteLine($"[Segment] Saved {Path.GetFileName(savePath)} ({compressedBytes.Length} bytes)"); } catch { }
                                // Truyền width và height của segment
                                try { Console.WriteLine($"[Segment] Debug - seg.Width: {seg.Width}, seg.Height: {seg.Height}"); } catch { }
                                TryPublishAnalyze(compressedBytes, seg.Width, seg.Height);
                            }
                            catch { }
                        }
                    }
                }
                try { Console.WriteLine($"[Segment] Done host={hostLabel}"); } catch { }
            }
            catch (Exception ex)
            {
                try { Console.WriteLine($"[Segment] ERROR {hostLabel}: {ex.Message}"); } catch { }
                ErrorWriter.WriteLog(LogPath, "SegmentScreenshots", ex.ToString());
                TelegramService.PushLogToTelegram($"SegmentScreenshots Error - {hostLabel}", ex);
            }
        }

        /// <summary>
        /// Đợi các quảng cáo trên trang web được render hoàn toàn trước khi chụp ảnh
        /// 
        /// Cách thức hoạt động:
        /// 1. Trong khoảng thời gian maxWait, liên tục kiểm tra:
        /// 2. Chạy JavaScript để tính điểm "score" dựa trên:
        ///    - iframe/frame: 
        ///      * Mỗi iframe có kích thước >= 120x30px => +2 điểm
        ///      * Nếu iframe có contentWindow => +1 điểm
        ///    - Các container quảng cáo (theo selector):
        ///      * Mỗi container có kích thước >= 120x30px => +2 điểm
        ///      * Nếu có thẻ <img> đã load hoàn toàn (complete && naturalWidth>0) => +2 điểm
        ///      * Nếu có thẻ <ins> với nội dung > 20 ký tự => +1 điểm
        ///    - Document readyState = "complete" => +1 điểm
        /// 3. Nếu score >= 4 và ổn định trong 2 lần kiểm tra liên tiếp => thoát (ads đã load xong)
        /// 4. Nếu chưa đủ điều kiện => đợi 350ms rồi kiểm tra lại
        /// 
        /// Mục đích: Đảm bảo tất cả quảng cáo đã render trước khi chụp để không bỏ sót
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - maxWait: Thời gian tối đa để đợi
        /// </summary>
        private static void WaitForAdsLoaded(IWebDriver driver, TimeSpan maxWait)
        {
            var js = (IJavaScriptExecutor)driver;
            var deadline = DateTime.UtcNow + maxWait;
            int stableRounds = 0;
            int lastScore = -1;
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    var script = @"
                        const minW=120, minH=30;
                        let score = 0;
                        // iframes likely ads
                        const iframes = Array.from(document.querySelectorAll('iframe, frame'));
                        for (const fr of iframes){
                          const r = fr.getBoundingClientRect();
                          if (r.width>=minW && r.height>=minH){ score += 2; }
                          try { if (fr.contentWindow) score += 1; } catch(e) {}
                        }
                        // common ad containers
                        const adSel = '[id*=''ad''],[id*=''ads''],[id*=''advert''],[id*=''banner''],[class*=''ad ''],[class*='' ad''],[class*=''ad-''],[class*=''-ad''],[class*=''ads''],[class*=''advert''],[class*=''banner''],.gpt-ad,.gpt-unit,.gpt-slot,.dfp-ad,.dfp-slot,.ad-slot,.ad-container,.ad-wrapper,.ad__container,.ad__slot,.adsbygoogle,.google-auto-placed,[data-ad],[data-ad-slot],[data-ad-unit],[data-google-query-id],[data-ez-name]';
                        const candidates = Array.from(document.querySelectorAll(adSel));
                        for (const el of candidates){
                          const r = el.getBoundingClientRect();
                          if (r.width>=minW && r.height>=minH) score += 2;
                          const img = el.querySelector('img');
                          if (img && img.complete && img.naturalWidth>0) score += 2;
                          const ins = el.querySelector('ins');
                          if (ins && ins.innerHTML && ins.innerHTML.trim().length>20) score += 1;
                        }
                        // readyState complete adds stability
                        if (document.readyState==='complete') score += 1;
                        return score;";
                    int score = Convert.ToInt32(js.ExecuteScript(script));
                    if (score == lastScore) stableRounds++; else { stableRounds = 0; lastScore = score; }
                    if (score >= 4 && stableRounds >= 2) break; // đủ tín hiệu rằng ads đã render ổn định
                }
                catch { }
                Thread.Sleep(350);
            }
        }

        /// <summary>
        /// Phủ lớp màu trắng lên các hình ảnh nhỏ trên trang để giảm dung lượng ảnh chụp màn hình
        /// 
        /// Cách thức hoạt động:
        /// 1. Tìm tất cả thẻ <img> trên document (không xử lý ảnh trong iframe)
        /// 2. Với mỗi ảnh:
        ///    - Lấy kích thước thực tế bằng getBoundingClientRect()
        ///    - Nếu kích thước nhỏ hơn minWidth HOẶC minHeight:
        ///      * Tạo một thẻ <div> với:
        ///        - position: fixed
        ///        - Vị trí và kích thước trùng khớp với ảnh
        ///        - background: màu trắng (#fff)
        ///        - z-index: tối đa để phủ lên trên
        ///        - pointer-events: none (không chặn click)
        ///      * Append vào document.body
        /// 3. Trả về số lượng overlay đã tạo
        /// 
        /// Mục đích: Các ảnh nhỏ (icon, avatar, thumbnail) thường không quan trọng cho việc phân tích quảng cáo
        /// Phủ trắng giúp nén ảnh tốt hơn, giảm dung lượng và tăng tốc độ xử lý
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - minWidth: Chiều rộng tối thiểu để không bị phủ (mặc định 120px)
        /// - minHeight: Chiều cao tối thiểu để không bị phủ (mặc định 80px)
        /// </summary>
        private static void MaskSmallImages(IWebDriver driver, int minWidth, int minHeight)
        {
            var js = (IJavaScriptExecutor)driver;
            var script = @"
                (function(){
                  try {
                    var overlays = 0;
                    var imgs = Array.from(document.images || []);
                    for (var i=0;i<imgs.length;i++){
                      var img = imgs[i];
                      var r = img.getBoundingClientRect();
                      if (r.width>0 && r.height>0 && (r.width < arguments[0] || r.height < arguments[1])){
                        var ov = document.createElement('div');
                        ov.style.position = 'fixed';
                        ov.style.left = r.left + 'px';
                        ov.style.top = r.top + 'px';
                        ov.style.width = r.width + 'px';
                        ov.style.height = r.height + 'px';
                        ov.style.background = '#fff';
                        ov.style.zIndex = '2147483647';
                        ov.style.pointerEvents = 'none';
                        document.body.appendChild(ov);
                        overlays++;
                      }
                    }
                    return overlays;
                  } catch(e){ return -1; }
                })();
            ";
            try { js.ExecuteScript(script, minWidth, minHeight); } catch { }
        }

        /// <summary>
        /// Kiểm tra link có hợp lệ không (phải chứa https:// hoặc http://)
        /// 
        /// Tham số:
        /// - link: Link cần kiểm tra
        /// 
        /// Trả về: true nếu link hợp lệ, false nếu không
        /// </summary>
        private static bool IsValidHttpLink(string link)
        {
            if (string.IsNullOrWhiteSpace(link))
            {
                return false;
            }
            
            // Loại bỏ khoảng trắng đầu cuối
            link = link.Trim();
            
            // Regex pattern để kiểm tra link HTTP/HTTPS hợp lệ
            // Phải bắt đầu với http:// hoặc https://
            var pattern = @"^https?://[^\s]+";
            var regex = new Regex(pattern, RegexOptions.IgnoreCase);
            
            return regex.IsMatch(link);
        }
        
        /// <summary>
        /// Normalize URL: loại bỏ fragment, normalize path
        /// </summary>
        private static string NormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return url;
            }
            
            try
            {
                var uri = new Uri(url);
                // Loại bỏ fragment (#...)
                var normalized = $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}{uri.Query}";
                return normalized;
            }
            catch
            {
                // Nếu không parse được URL, trả về nguyên bản
                return url.Trim();
            }
        }
        
        /// <summary>
        /// Thu thập tất cả links trên trang, phân loại quảng cáo và push vào RabbitMQ
        /// 
        /// Cách thức hoạt động:
        /// 1. Thu thập tất cả links (<a href>) trên trang
        /// 2. Phân tích từng link để xác định có phải quảng cáo không:
        ///    - Phân tích DOM: class/id chứa "ad", text "Quảng cáo", iframe ads, banner images
        ///    - Phân tích Network: domain quảng cáo (doubleclick.net, googlesyndication.com, ...)
        ///    - Heuristic: URL pattern (utm_source, adclick, ...), vị trí, kích thước
        /// 3. Push các link quảng cáo vào queue RabbitQueueAnalyzeSingleBanner với JSON format:
        ///    { "link_click_banner": "{link}", "screenshot_base64": "" }
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - hostLabel: Hostname để log
        /// </summary>
        public static void ProcessIframesAndPushToQueue(IWebDriver driver, string hostLabel)
        {
            try
            {
                // Kiểm tra nếu RabbitQueueScreenLink chứa "NO_RUN" thì bỏ qua hàm này
                if (!string.IsNullOrWhiteSpace(RabbitQueueScreenLink) && RabbitQueueScreenLink.Contains("NO_RUN", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("[AdLink] RabbitQueueScreenLink có chứa 'NO_RUN', bỏ qua ProcessIframesAndPushToQueue.");
                    return;
                }
                Console.WriteLine($"[AdLink] Bắt đầu thu thập links từ HTML source cho host={hostLabel}");
                
                // Lấy HTML source từ driver
                var htmlSource = driver.PageSource;
                Console.WriteLine($"[AdLink] Debug - HTML source length: {htmlSource?.Length ?? 0} characters");
                
                if (string.IsNullOrWhiteSpace(htmlSource))
                {
                    Console.WriteLine($"[AdLink] HTML source rỗng, không thể parse links");
                    return;
                }
                
                // Parse links từ HTML bằng C#
                var allLinks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                
                // 1. Tìm tất cả href trong <a> tags với target="_blank"
                var hrefPattern = @"<a(?=[^>]*\btarget\s*=\s*[""']?_blank[""']?)[^>]+href\s*=\s*[""']([^""']+)[""'][^>]*>";
                var hrefMatches = Regex.Matches(htmlSource, hrefPattern, RegexOptions.IgnoreCase);
                Console.WriteLine($"[AdLink] Debug - Tìm thấy {hrefMatches.Count} <a href target=\"_blank\"> tags");
                foreach (Match match in hrefMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        var href = match.Groups[1].Value.Trim();
                        if (IsValidHttpLink(href))
                        {
                            allLinks.Add(NormalizeUrl(href));
                        }
                    }
                }
                
                // 2. Tìm tất cả src trong <iframe> và <frame> tags
                var iframePattern = @"<(?:iframe|frame)[^>]+src\s*=\s*[""']([^""']+)[""'][^>]*>";
                var iframeMatches = Regex.Matches(htmlSource, iframePattern, RegexOptions.IgnoreCase);
                Console.WriteLine($"[AdLink] Debug - Tìm thấy {iframeMatches.Count} <iframe/frame src> tags");
                foreach (Match match in iframeMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        var src = match.Groups[1].Value.Trim();
                        if (IsValidHttpLink(src))
                        {
                            allLinks.Add(NormalizeUrl(src));
                        }
                    }
                }
                
                // 3. Tìm data-href và data-url
                var dataHrefPattern = @"data-href\s*=\s*[""']([^""']+)[""']";
                var dataHrefMatches = Regex.Matches(htmlSource, dataHrefPattern, RegexOptions.IgnoreCase);
                Console.WriteLine($"[AdLink] Debug - Tìm thấy {dataHrefMatches.Count} data-href attributes");
                foreach (Match match in dataHrefMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        var href = match.Groups[1].Value.Trim();
                        if (IsValidHttpLink(href))
                        {
                            allLinks.Add(NormalizeUrl(href));
                        }
                    }
                }
                
                var dataUrlPattern = @"data-url\s*=\s*[""']([^""']+)[""']";
                var dataUrlMatches = Regex.Matches(htmlSource, dataUrlPattern, RegexOptions.IgnoreCase);
                Console.WriteLine($"[AdLink] Debug - Tìm thấy {dataUrlMatches.Count} data-url attributes");
                foreach (Match match in dataUrlMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        var url = match.Groups[1].Value.Trim();
                        if (IsValidHttpLink(url))
                        {
                            allLinks.Add(NormalizeUrl(url));
                        }
                    }
                }
                
                // 4. Tìm URLs trong onclick handlers
                var onclickPattern = @"onclick\s*=\s*[""']([^""']+)[""']";
                var onclickMatches = Regex.Matches(htmlSource, onclickPattern, RegexOptions.IgnoreCase);
                Console.WriteLine($"[AdLink] Debug - Tìm thấy {onclickMatches.Count} onclick handlers");
                foreach (Match match in onclickMatches)
                {
                    if (match.Groups.Count > 1)
                    {
                        var onclickContent = match.Groups[1].Value;
                        // Tìm URLs trong onclick content
                        var urlInOnclick = Regex.Matches(onclickContent, @"https?://[^\s'""\)]+", RegexOptions.IgnoreCase);
                        foreach (Match urlMatch in urlInOnclick)
                        {
                            var url = urlMatch.Value.Trim();
                            if (IsValidHttpLink(url))
                            {
                                allLinks.Add(NormalizeUrl(url));
                            }
                        }
                    }
                }
                
                Console.WriteLine($"[AdLink] Tổng cộng thu thập được {allLinks.Count} link(s) HTTP/HTTPS hợp lệ");
                
                // Phân loại link quảng cáo bằng C#
                var adDomains = new[]
                {
                    "doubleclick.net", "googlesyndication.com", "adservice.google.com",
                    "adclick.g.doubleclick.net", "googleadservices.com",
                    "adsystem.amazon.com", "amazon-adsystem.com",
                    "ads.yahoo.com", "taboola.com", "outbrain.com", "adnxs.com",
                    "criteo.com", "adsrvr.org", "advertising.com", "adtech.com",
                    "adform.net", "rubiconproject.com", "openx.net", "adsafeprotected.com",
                };
                
                string currentDomain = string.Empty;
                try { currentDomain = new Uri(driver.Url).Host; } catch { currentDomain = hostLabel ?? string.Empty; }
                string CleanHost(string h) => string.IsNullOrWhiteSpace(h) ? string.Empty : h.Replace("www.", string.Empty, StringComparison.OrdinalIgnoreCase);
                currentDomain = CleanHost(currentDomain);
                
                int ClassifyScore(string href)
                {
                    int score = 0;
                    try
                    {
                        var uri = new Uri(href);
                        var host = CleanHost(uri.Host).ToLowerInvariant();
                        // Ad domain exact/subdomain
                        foreach (var ad in adDomains)
                        {
                            if (host == ad || host.EndsWith("." + ad, StringComparison.OrdinalIgnoreCase))
                            {
                                score += 5; break;
                            }
                        }
                        // URL patterns
                        var urlLower = href.ToLowerInvariant();
                        if (urlLower.Contains("adclick") || urlLower.Contains("utm_source=ad") ||
                            urlLower.Contains("utm_medium=ad") || urlLower.Contains("click?") ||
                            urlLower.Contains("impression") || urlLower.Contains("gclid") ||
                            urlLower.Contains("fbclid") || urlLower.Contains("ref=ad") ||
                            urlLower.Contains("promo") || urlLower.Contains("sponsor"))
                        {
                            score += 3;
                        }
                        // Path hints
                        if (uri.AbsolutePath.Contains("banner", StringComparison.OrdinalIgnoreCase) ||
                            uri.AbsolutePath.Contains("ad", StringComparison.OrdinalIgnoreCase))
                        {
                            score += 1;
                        }
                        // External domain
                        var linkDomain = host;
                        if (!string.IsNullOrEmpty(currentDomain) &&
                            !linkDomain.Equals(currentDomain, StringComparison.OrdinalIgnoreCase) &&
                            !linkDomain.Contains(currentDomain, StringComparison.OrdinalIgnoreCase) &&
                            !currentDomain.Contains(linkDomain, StringComparison.OrdinalIgnoreCase))
                        {
                            score += 1;
                        }
                    }
                    catch { }
                    return score;
                }
                
                var adOnly = new List<(string href,int score)>();
                foreach (var href in allLinks)
                {
                    int s = ClassifyScore(href);
                    if (s >= 2) adOnly.Add((href,s));
                }
                
                Console.WriteLine($"[AdLink] Sau phân loại: {adOnly.Count} / {allLinks.Count} link có khả năng quảng cáo (>=2 điểm)");
                
                // Lấy job params từ RabbitQueue gốc để merge vào payload
                var jobParamsSnapshot = lastJobParams; // tránh race condition
                
                // Push chỉ link quảng cáo
                int processedCount = 0;
                foreach (var item in adOnly)
                {
                    var href = item.href;
                    try
                    {
                        PushIframeToQueue(href, "", jobParamsSnapshot);
                        processedCount++;
                        Console.WriteLine($"[AdLink] Pushed {processedCount}/{adOnly.Count} - score={item.score} - {href.Substring(0, Math.Min(80, href.Length))}...");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[AdLink] Lỗi khi push link: {ex.Message}");
                        ErrorWriter.WriteLog(LogPath, "ProcessAdLink", $"{hostLabel} => {href} => {ex}");
                    }
                }
                
                Console.WriteLine($"[AdLink] Hoàn thành xử lý {processedCount} link(s) quảng cáo cho host={hostLabel}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] Lỗi xử lý links: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "ProcessIframesAndPushToQueue", $"{hostLabel} => {ex}");
                TelegramService.PushLogToTelegram($"ProcessIframesAndPushToQueue Error - {hostLabel}", ex);
            }
        }
        
        /// <summary>
        /// Push link iframe vào RabbitMQ queue RabbitQueueAnalyzeSingleBanner
        /// 
        /// Cách thức hoạt động:
        /// 1. Kiểm tra queue đã được cấu hình
        /// 2. Tạo JSON payload: 
        ///    - Merge tất cả thông tin từ job params gốc (link_web, slice, quanlity_image, device, ...)
        ///    - Thêm "link_click_banner": "{link}"
        ///    - Thêm "screenshot_base64": "" (luôn rỗng)
        /// 3. Đảm bảo kết nối RabbitMQ đã sẵn sàng
        /// 4. Push message lên queue với persistent = true
        /// 
        /// Tham số:
        /// - linkClick: Link click của iframe/quảng cáo
        /// - screenshotBase64: Luôn để rỗng (không chụp ảnh nữa)
        /// - jobParamsSnapshot: Thông tin từ job gốc (RabbitQueue) để merge vào payload
        /// </summary>
        private static void PushIframeToQueue(string linkClick, string screenshotBase64, JObject? jobParamsSnapshot = null)
        {
            try
            {
            
                // Gửi link tới queue screen link để job chụp ảnh từ link click xử lý
                var targetQueue = RabbitQueueScreenLink; // Queue name này dùng để nhận dữ liệu từ link click banner
                
                if (string.IsNullOrWhiteSpace(linkClick))
                {
                    Console.WriteLine("[Iframe] Link click rỗng, bỏ qua");
                    return;
                }
                
                // Tạo JSON payload: merge tất cả thông tin từ job params gốc
                JObject mergedPayload;
                if (jobParamsSnapshot != null)
                {
                    try 
                    { 
                        mergedPayload = (JObject)jobParamsSnapshot.DeepClone(); 
                    }
                    catch 
                    { 
                        mergedPayload = new JObject(); 
                    }
                }
                else
                {
                    mergedPayload = new JObject();
                }
                
                // Thêm link_click_banner và screenshot_base64 (ghi đè nếu có trong job params)
                mergedPayload["link_click_banner"] = linkClick ?? "";
                mergedPayload["screenshot_base64"] = ""; // Luôn để rỗng
                
                var jsonPayload = mergedPayload.ToString(Newtonsoft.Json.Formatting.None);
                var body = Encoding.UTF8.GetBytes(jsonPayload);
                
                // Debug: Log payload để kiểm tra
                try
                {
                    var payloadKeys = string.Join(", ", mergedPayload.Properties().Select(p => p.Name));
                    Console.WriteLine($"[Iframe] Debug - Payload keys: {payloadKeys}");
                }
                catch { }
                
                // Đảm bảo kết nối RabbitMQ sẵn sàng với queue name từ config
                EnsureIframePublisherReady(targetQueue);
                
                if (iframeChannel == null || !iframeChannel.IsOpen)
                {
                    Console.WriteLine("[Iframe] Không thể kết nối tới RabbitMQ");
                    return;
                }
                
                var props = iframeChannel.CreateBasicProperties();
                props.Persistent = true;
                
                iframeChannel.BasicPublish(exchange: "",
                                          routingKey: targetQueue,
                                          basicProperties: props,
                                          body: body);
                
                Console.WriteLine($"[Iframe] Đã push vào queue '{targetQueue}' - Link: {linkClick}, Size: {body.Length} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Iframe] Lỗi push vào queue: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "PushIframeToQueue", ex.ToString());
            }
        }
        
        // Publisher for iframe queue
        private static readonly object iframePubLock = new object();
        private static IConnection? iframeConnection;
        private static IModel? iframeChannel;
        
        /// <summary>
        /// Đảm bảo kết nối RabbitMQ cho iframe publisher đã sẵn sàng
        /// </summary>
        private static void EnsureIframePublisherReady(string queueName)
        {
            if (iframeChannel != null && iframeChannel.IsOpen) return;
            
            lock (iframePubLock)
            {
                try
                {
                    if (iframeChannel != null && iframeChannel.IsOpen) return;
                    
                    iframeConnection?.Dispose();
                    iframeChannel?.Dispose();
                    
                    var factory = new ConnectionFactory()
                    {
                        HostName = rabbit_host,
                        UserName = rabbit_username,
                        Password = rabbit_password,
                        VirtualHost = string.IsNullOrWhiteSpace(rabbit_vhost) ? "/" : rabbit_vhost,
                        Port = rabbit_port,
                        AutomaticRecoveryEnabled = true,
                        NetworkRecoveryInterval = TimeSpan.FromSeconds(5),
                        RequestedConnectionTimeout = TimeSpan.FromSeconds(30),
                        RequestedHeartbeat = TimeSpan.FromSeconds(30)
                    };
                    
                    if (rabbit_use_ssl == "1" || rabbit_port == 5671)
                    {
                        factory.Ssl = new SslOption
                        {
                            Enabled = true,
                            ServerName = rabbit_host,
                            AcceptablePolicyErrors = System.Net.Security.SslPolicyErrors.None
                        };
                    }
                    
                    iframeConnection = factory.CreateConnection();
                    iframeChannel = iframeConnection.CreateModel();
                    iframeChannel.QueueDeclare(queue: queueName,
                                              durable: true,
                                              exclusive: false,
                                              autoDelete: false,
                                              arguments: null);
                    
                    Console.WriteLine($"[Iframe] Đã kết nối RabbitMQ và declare queue: '{queueName}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Iframe] Lỗi kết nối RabbitMQ: {ex.Message}");
                    ErrorWriter.WriteLog(LogPath, "EnsureIframePublisher", ex.ToString());
                }
            }
        }

        /// <summary>
        /// Gửi ảnh đã chụp lên RabbitMQ queue để AI (Gemini) phân tích
        /// 
        /// Cách thức hoạt động:
        /// 1. Kiểm tra ảnh hợp lệ (không null, không rỗng)
        /// 2. Kiểm tra queue analyze đã được cấu hình
        /// 3. Đảm bảo kết nối RabbitMQ analyze publisher đã sẵn sàng (EnsureAnalyzePublisherReady)
        /// 4. Build payload chứa:
        ///    - Ảnh (dạng base64 hoặc raw bytes tùy theo analyze_publish_raw)
        ///    - Thông tin từ job gốc (link_web, slice, quanlity_image, ...)
        ///    - width và height của ảnh
        /// 5. Gửi message lên queue RabbitQueueAnalyze với persistent = true (không mất khi server restart)
        /// 6. Log kích thước message đã gửi
        /// 
        /// Nếu lỗi: Ghi log nhưng không throw exception (tránh ảnh hưởng đến luồng chính)
        /// 
        /// Tham số:
        /// - imageBytes: Mảng byte chứa dữ liệu ảnh đã nén
        /// - width: Chiều rộng của ảnh (optional)
        /// - height: Chiều cao của ảnh (optional)
        /// </summary>
        public static void TryPublishAnalyze(byte[] imageBytes, int? width = null, int? height = null)
        {
            try
            {
                if (imageBytes == null || imageBytes.Length == 0) return;
                if (string.IsNullOrWhiteSpace(RabbitQueueAnalyze)) return;
                EnsureAnalyzePublisherReady();
                if (analyzeChannel == null) return;

                var snapshot = lastJobParams; // tránh race
                
                // Debug: Log width và height trước khi build payload
                try
                {
                    Console.WriteLine($"[TryPublishAnalyze] Debug - width: {(width.HasValue ? width.Value.ToString() : "null")}, height: {(height.HasValue ? height.Value.ToString() : "null")}");
                }
                catch { }
                
                var body = AnalyzePayloadBuilder.BuildAnalyzeBody(imageBytes, snapshot, analyze_publish_raw, width, height);

                var props = analyzeChannel.CreateBasicProperties();
                props.Persistent = true;

                analyzeChannel.BasicPublish(exchange: "",
                                            routingKey: RabbitQueueAnalyze,
                                            basicProperties: props,
                                            body: body);
                
                // Log thông tin payload
                if (analyze_publish_raw != "1" && width.HasValue && height.HasValue)
                {
                    try 
                    { 
                        Console.WriteLine($"[Analyze] Published {body.Length} bytes to queue '{RabbitQueueAnalyze}' - Width: {width.Value}, Height: {height.Value}"); 
                    } 
                    catch { }
                }
                else
                {
                    try { Console.WriteLine($"[Analyze] Published {(analyze_publish_raw=="1"?imageBytes.Length:body.Length)} bytes to queue '{RabbitQueueAnalyze}' (raw={analyze_publish_raw})"); } catch { }
                }
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "PublishAnalyze", ex.ToString());
            }
        }

        /// <summary>
        /// Đảm bảo kết nối RabbitMQ cho analyze publisher đã sẵn sàng
        /// 
        /// Cách thức hoạt động:
        /// 1. Kiểm tra channel đã tồn tại và đang mở => return luôn
        /// 2. Sử dụng lock để tránh race condition khi nhiều thread cùng tạo connection
        /// 3. Double-check: Kiểm tra lại channel sau khi vào lock
        /// 4. Dispose connection và channel cũ nếu có
        /// 5. Tạo ConnectionFactory với cấu hình:
        ///    - Host, User, Password, VirtualHost, Port từ config
        ///    - Bật AutomaticRecoveryEnabled để tự động reconnect khi mất kết nối
        ///    - NetworkRecoveryInterval: 5 giây
        ///    - RequestedConnectionTimeout: 30 giây
        ///    - RequestedHeartbeat: 30 giây
        ///    - Nếu rabbit_use_ssl = "1" hoặc port = 5671 => bật SSL
        /// 6. Tạo connection và channel mới
        /// 7. Declare queue RabbitQueueAnalyze với durable = true (không mất khi server restart)
        /// 
        /// Nếu lỗi: Ghi log nhưng không throw (graceful degradation)
        /// </summary>
        private static void EnsureAnalyzePublisherReady()
        {
            if (analyzeChannel != null && analyzeChannel.IsOpen) return;
            lock (analyzePubLock)
            {
                try
                {
                    if (analyzeChannel != null && analyzeChannel.IsOpen) return;

                    analyzeConnection?.Dispose();
                    analyzeChannel?.Dispose();

                    var factory = new ConnectionFactory()
                    {
                        HostName = rabbit_host,
                        UserName = rabbit_username,
                        Password = rabbit_password,
                        VirtualHost = string.IsNullOrWhiteSpace(rabbit_vhost) ? "/" : rabbit_vhost,
                        Port = rabbit_port,
                        AutomaticRecoveryEnabled = true,
                        NetworkRecoveryInterval = TimeSpan.FromSeconds(5),
                        RequestedConnectionTimeout = TimeSpan.FromSeconds(30),
                        RequestedHeartbeat = TimeSpan.FromSeconds(30)
                    };
                    if (rabbit_use_ssl == "1" || rabbit_port == 5671)
                    {
                        factory.Ssl = new SslOption
                        {
                            Enabled = true,
                            ServerName = rabbit_host,
                            AcceptablePolicyErrors = System.Net.Security.SslPolicyErrors.None
                        };
                    }

                    analyzeConnection = factory.CreateConnection();
                    analyzeChannel = analyzeConnection.CreateModel();
                    analyzeChannel.QueueDeclare(queue: RabbitQueueAnalyze,
                                                durable: true,
                                                exclusive: false,
                                                autoDelete: false,
                                                arguments: null);
                }
                catch (Exception ex)
                {
                    ErrorWriter.WriteLog(LogPath, "EnsureAnalyzePublisher", ex.ToString());
                }
            }
        }

        /// <summary>
        /// Phát hiện các vùng chứa quảng cáo trên trang và trả về danh sách tọa độ (top, bottom)
        /// 
        /// Cách thức hoạt động:
        /// 1. Lấy danh sách selector quảng cáo phổ biến (GetCommonAdSelectors)
        /// 2. Chạy JavaScript để:
        ///    - Tìm tất cả element khớp với selector
        ///    - Lấy tọa độ của mỗi element bằng getBoundingClientRect()
        ///    - Chỉ tính các element có kích thước >= 120x30px (kích thước quảng cáo tối thiểu)
        ///    - Tính top/bottom theo tọa độ tài liệu (scrollY + getBoundingClientRect().top)
        ///    - Trả về mảng [[top1, bottom1], [top2, bottom2], ...]
        /// 3. Parse kết quả từ JavaScript về C# List<(int Top, int Bottom)>
        /// 4. Gộp các rect chồng lấn bằng MergeOverlappingRects để tối ưu
        /// 
        /// Mục đích: Xác định vị trí quảng cáo để tránh cắt ảnh ngang qua quảng cáo khi chia segment
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// 
        /// Trả về: Danh sách (Top, Bottom) của các vùng quảng cáo, đã được gộp nếu chồng lấn
        /// </summary>
        private static List<(int Top, int Bottom)> TryDetectAdRects(IWebDriver driver)
        {
            var results = new List<(int Top, int Bottom)>();
            try
            {
                var selectors = string.Join(", ", GetCommonAdSelectors());
                var script = @"
                    const sels = arguments[0];
                    const minW = 120, minH = 30;
                    const list = Array.from(document.querySelectorAll(sels));
                    const y = window.scrollY || window.pageYOffset || 0;
                    const rects = [];
                    for (const el of list) {
                        try {
                            const r = el.getBoundingClientRect();
                            const w = Math.round(r.width);
                            const h = Math.round(r.height);
                            if (w >= minW && h >= minH) {
                                const top = Math.max(0, Math.round(r.top + y));
                                const bottom = Math.max(top, Math.round(r.bottom + y));
                                rects.push([top, bottom]);
                            }
                        } catch {}
                    }
                    rects
                ";
                var raw = (System.Collections.IEnumerable)((IJavaScriptExecutor)driver).ExecuteScript(script, selectors);
                foreach (var item in raw)
                {
                    var pair = item as System.Collections.IList;
                    if (pair != null && pair.Count >= 2)
                    {
                        int top = Convert.ToInt32(pair[0]);
                        int bottom = Convert.ToInt32(pair[1]);
                        if (bottom > top) results.Add((top, bottom));
                    }
                }

                // Gộp/đơn giản hoá các rect chồng lấn để tính toán nhanh hơn
                results = MergeOverlappingRects(results);
            }
            catch { }
            return results;
        }

        /// <summary>
        /// Gộp các khoảng tọa độ [top, bottom] nếu chúng chồng lấn hoặc gần nhau
        /// 
        /// Cách thức hoạt động:
        /// 1. Nếu danh sách rỗng => trả về list rỗng
        /// 2. Sắp xếp các rect theo Top (từ trên xuống)
        /// 3. Duyệt qua từng rect:
        ///    - Lấy rect đầu tiên làm current
        ///    - Với mỗi rect tiếp theo:
        ///      * Nếu Top của rect tiếp theo <= Bottom của current (chồng lấn):
        ///        - Gộp: current.Bottom = max(current.Bottom, rect.Bottom)
        ///      * Nếu không chồng lấn:
        ///        - Thêm current vào kết quả
        ///        - Set current = rect tiếp theo
        /// 4. Thêm rect cuối cùng vào kết quả
        /// 
        /// Ví dụ: [(0,100), (50,150), (200,300)] => [(0,150), (200,300)]
        /// 
        /// Mục đích: Giảm số lượng rect để tối ưu tính toán sau này
        /// 
        /// Tham số:
        /// - rects: Danh sách các khoảng (Top, Bottom) có thể chồng lấn
        /// 
        /// Trả về: Danh sách đã được gộp, không còn chồng lấn
        /// </summary>
        private static List<(int Top, int Bottom)> MergeOverlappingRects(List<(int Top, int Bottom)> rects)
        {
            if (rects == null || rects.Count == 0) return new List<(int Top, int Bottom)>();
            var ordered = rects.OrderBy(r => r.Top).ToList();
            var merged = new List<(int Top, int Bottom)>();
            var cur = ordered[0];
            for (int i = 1; i < ordered.Count; i++)
            {
                var r = ordered[i];
                if (r.Top <= cur.Bottom)
                {
                    cur = (cur.Top, Math.Max(cur.Bottom, r.Bottom));
                }
                else
                {
                    merged.Add(cur);
                    cur = r;
                }
            }
            merged.Add(cur);
            return merged;
        }

        /// <summary>
        /// Điều chỉnh vị trí bắt đầu của segment để tránh cắt ngang qua quảng cáo
        /// 
        /// Cách thức hoạt động:
        /// 1. Nếu không có quảng cáo => trả về proposedY (có giới hạn trong phạm vi trang)
        /// 2. Kiểm tra biên trên và biên dưới của viewport có an toàn không:
        ///    - Biên trên: y + margin
        ///    - Biên dưới: y + viewportHeight - margin
        ///    - An toàn = không nằm giữa bất kỳ quảng cáo nào
        /// 3. Nếu vị trí đề xuất (proposedY) an toàn => trả về luôn
        /// 4. Nếu không an toàn => tìm vị trí gần nhất an toàn:
        ///    - Tìm trong bán kính searchRadius (min(200, viewportHeight/3))
        ///    - Thử từ proposedY ± 0, 1, 2, ... pixels
        ///    - Dừng khi tìm thấy vị trí an toàn
        /// 5. Nếu không tìm được => điều chỉnh đến mép quảng cáo gần nhất:
        ///    - Nếu biên trên nằm giữa quảng cáo => đưa lên top của quảng cáo - margin
        ///    - Nếu biên dưới nằm giữa quảng cáo => đẩy xuống bottom của quảng cáo + margin
        /// 6. Đảm bảo kết quả >= minY và <= totalHeight - viewportHeight
        /// 
        /// Mục đích: Tránh cắt ảnh ngang qua quảng cáo để AI dễ nhận diện quảng cáo nguyên vẹn
        /// 
        /// Tham số:
        /// - proposedY: Vị trí Y đề xuất ban đầu
        /// - viewportHeight: Chiều cao viewport để chụp
        /// - totalHeight: Chiều cao tổng của trang
        /// - adRects: Danh sách vùng quảng cáo đã được gộp
        /// - minY: Vị trí Y tối thiểu cho phép
        /// 
        /// Trả về: Vị trí Y đã được điều chỉnh để tránh quảng cáo
        /// </summary>
        private static int AdjustSliceStartToAvoidAds(int proposedY, int viewportHeight, int totalHeight, List<(int Top, int Bottom)> adRects, int minY)
        {
            if (adRects == null || adRects.Count == 0)
            {
                return Math.Max(0, Math.Min(proposedY, totalHeight - viewportHeight));
            }

            int margin = 4; // khoảng đệm nhỏ
            int bestY = proposedY;

            Func<int, bool> isBoundarySafe = (y) =>
            {
                int topLine = y + margin;
                int bottomLine = y + viewportHeight - margin;
                foreach (var r in adRects)
                {
                    if (r.Top < topLine && topLine < r.Bottom) return false;      // biên trên cắt ngang
                    if (r.Top < bottomLine && bottomLine < r.Bottom) return false; // biên dưới cắt ngang
                }
                return true;
            };

            bestY = Math.Max(minY, Math.Max(0, Math.Min(proposedY, totalHeight - viewportHeight)));
            if (isBoundarySafe(bestY)) return bestY;

            // Thử dịch lên/xuống trong một khoảng giới hạn để tìm điểm an toàn gần nhất
            int searchRadius = Math.Min(200, viewportHeight / 3);
            int lowBound = Math.Max(minY, Math.Max(0, bestY - searchRadius));
            int highBound = Math.Min(totalHeight - viewportHeight, bestY + searchRadius);

            int nearest = bestY;
            int bestDist = int.MaxValue;
            for (int dy = 0; dy <= searchRadius; dy++)
            {
                int up = bestY - dy;
                if (up >= lowBound && isBoundarySafe(up))
                {
                    nearest = up; bestDist = dy; break;
                }
                int down = bestY + dy;
                if (down <= highBound && isBoundarySafe(down))
                {
                    nearest = down; bestDist = dy; break;
                }
            }

            if (bestDist != int.MaxValue) return nearest;

            // Nếu không tìm được điểm hoàn hảo, rơi về việc neo tại mép của khối gần nhất (ưu tiên neo trên/dưới)
            foreach (var r in adRects)
            {
                // nếu biên trên rơi giữa khối, đưa lên đầu khối
                if (r.Top < bestY && bestY < r.Bottom)
                {
                    int candidate = Math.Max(minY, Math.Min(r.Top - margin, totalHeight - viewportHeight));
                    if (candidate >= minY) return candidate;
                }
                // nếu biên dưới rơi giữa khối, đẩy xuống cuối khối
                int bottomLine = bestY + viewportHeight;
                if (r.Top < bottomLine && bottomLine < r.Bottom)
                {
                    int candidate = Math.Min(totalHeight - viewportHeight, r.Bottom + margin);
                    if (candidate >= minY) return candidate;
                }
            }

            return Math.Max(minY, Math.Max(0, Math.Min(bestY, totalHeight - viewportHeight)));
        }

        /// <summary>
        /// Xử lý đặc biệt cho website ThanhNien.vn để chụp quảng cáo
        /// 
        /// Cách thức hoạt động:
        /// 1. Đợi trang web load hoàn toàn (readyState = "complete")
        /// 2. Đảm bảo các quảng cáo ở phần đầu trang đã được render (EnsureTopAdsNearTop)
        /// 3. Tạo danh sách selector quảng cáo đặc thù cho ThanhNien:
        ///    - Selector chung từ GetCommonAdSelectors()
        ///    - Selector riêng: header .banner, #banner_top, .gpt-ad, .qc, .quangcao
        /// 4. Chụp các element quảng cáo bằng CaptureBySelectors
        ///    - Tìm các element khớp với selector
        ///    - Cuộn element vào view
        ///    - Đợi element render
        ///    - Cắt ảnh từ screenshot toàn trang và lưu
        /// 5. Chụp các iframe quảng cáo bằng CaptureAdIframes
        ///    - Tìm tất cả iframe/frame
        ///    - Lọc iframe có kích thước >= 120x30px
        ///    - Chụp và lưu từng iframe
        /// 
        /// Tham số:
        /// - driver: WebDriver đã điều hướng tới ThanhNien
        /// - jpegQuality: Chất lượng JPEG để nén ảnh
        /// </summary>
        private static void HandleThanhNien(IWebDriver driver, long jpegQuality)
        {
            var wait = new WebDriverWait(new SystemClock(), driver, TimeSpan.FromSeconds(15), TimeSpan.FromMilliseconds(250));
            try
            {
                wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").ToString() == "complete");
            }
            catch { }

            // Đảm bảo khu vực top ads sẵn sàng
            EnsureTopAdsNearTop(driver, TimeSpan.FromSeconds(8));

            var siteSpecific = new[]
            {
                "header .banner, header [id*='banner'], header [class*='banner']",
                "#banner_top, .banner-top, .top-banner, .leaderboard, .top-ads, .banner-leaderboard",
                "[id^='div-gpt-ad'], [id*='gpt-ad'], .gpt-ad, .dfp-ad, .ad-slot, .ad-container",
                ".qc, .quangcao"
            };
            var selectors = AdCapture.GetCommonAdSelectors().Concat(siteSpecific).ToArray();
            AdCapture.CaptureBySelectors(driver, selectors, "thanhnien.vn", startupPath, LogPath, jpegQuality);
            AdCapture.CaptureAdIframes(driver, "thanhnien.vn", startupPath, LogPath, jpegQuality);
        }

        /// <summary>
        /// Chụp quảng cáo chung cho các website không có handler riêng
        /// 
        /// Cách thức hoạt động:
        /// 1. Đảm bảo quảng cáo ở phần đầu trang đã render (EnsureTopAdsNearTop)
        /// 2. Tạo danh sách selector quảng cáo generic:
        ///    - Selector chung từ GetCommonAdSelectors() (gpt-ad, dfp-ad, adsbygoogle, ...)
        ///    - Selector generic: div.banner, [id*='banner'], [class*='banner-ads'], ...
        /// 3. Chụp các element quảng cáo:
        ///    - Duyệt qua từng selector
        ///    - Tìm tối đa 10 element đầu tiên
        ///    - Cuộn element vào view, đợi render
        ///    - Cắt ảnh và lưu vào screenshots/<hostLabel>/
        ///    - Tự động publish lên analyze queue
        /// 4. Chụp các iframe quảng cáo (nhiều quảng cáo dùng iframe)
        /// 
        /// Ưu điểm: Áp dụng được cho hầu hết website không cần xử lý đặc biệt
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - hostLabel: Hostname để đặt tên thư mục
        /// - jpegQuality: Chất lượng JPEG
        /// </summary>
        private static void CaptureGenericBanners(IWebDriver driver, string hostLabel, long jpegQuality)
        {
            // Đảm bảo khu vực top ads sẵn sàng trên các trang khác
            EnsureTopAdsNearTop(driver, TimeSpan.FromSeconds(6));

            var siteSpecific = new[]
            {
                "div.banner, section.banner, header .banner",
                "[id*='banner'], [class*='banner'], .banner-ads, [class*='banner-ads']",
                "[id*='ads'], [class*='ads'], [class*='ad-'], [class*='advert']"
            };
            var selectors = AdCapture.GetCommonAdSelectors().Concat(siteSpecific).ToArray();
            // Chụp các phần tử quảng cáo trong DOM chính
            AdCapture.CaptureBySelectors(driver, selectors, hostLabel, startupPath, LogPath, jpegQuality);
            // Chụp trực tiếp các iframe
            AdCapture.CaptureAdIframes(driver, hostLabel, startupPath, LogPath, jpegQuality);
        }

        /// <summary>
        /// Wrapper function để giữ tương thích với code cũ
        /// 
        /// Cách thức hoạt động:
        /// - Chuyển giao toàn bộ xử lý sang Utils.AdCapture.CaptureBySelectors
        /// - Truyền thêm tham số jpegQuality để nén ảnh
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - selectors: Mảng CSS selector để tìm quảng cáo
        /// - hostLabel: Hostname
        /// - jpegQuality: Chất lượng JPEG
        /// </summary>
        private static void CaptureBySelectors(IWebDriver driver, string[] selectors, string hostLabel, long jpegQuality)
        {
            AdCapture.CaptureBySelectors(driver, selectors, hostLabel, startupPath, LogPath, jpegQuality);
        }
     
        /// <summary>
        /// Lưu lại ảnh chụp màn hình của một phần tử (element), cắt từ screenshot toàn trang, vào thư mục theo hostLabel.
        /// Nếu không cắt được (tọa độ vượt ngoài ảnh), sẽ lưu toàn bộ screenshot thay thế.
        /// </summary>
        // (moved to Utils/AdCapture.cs)

        // Tạo nhãn tên phần tử từ id hoặc class để đặt tên file     

        /// <summary>
        /// Đảm bảo các quảng cáo ở phần đầu trang đã được render (hàm generic áp dụng cho mọi website)
        /// 
        /// Cách thức hoạt động:
        /// 1. Cuộn lên đầu trang (scrollTo(0,0))
        /// 2. Trong khoảng thời gian maxWait, liên tục:
        ///    - Lắc scroll nhẹ (scroll 60px xuống rồi quay lại 0) để kích hoạt lazy-loading
        ///    - Chạy JavaScript để đếm số lượng quảng cáo hiển thị:
        ///      * Tìm các selector: header banner, #banner_top, .gpt-ad, iframe, ...
        ///      * Chỉ tính quảng cáo có kích thước >= 120x30px và nằm trong vùng top (y < 900px)
        ///    - Nếu số lượng quảng cáo = 0 => đợi 220ms rồi kiểm tra lại
        ///    - Nếu số lượng ổn định trong 2 lần kiểm tra liên tiếp => thoát
        ///    - Nếu số lượng thay đổi => reset counter và tiếp tục
        /// 3. Thoát khi đã đợi hết maxWait hoặc quảng cáo đã ổn định
        /// 
        /// Khác với EnsureTopAdsOnVnExpress: Hàm này generic, có thể dùng cho mọi website
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - maxWait: Thời gian tối đa để đợi
        /// </summary>
        private static void EnsureTopAdsNearTop(IWebDriver driver, TimeSpan maxWait)
        {
            var js = (IJavaScriptExecutor)driver;
            try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }

            var end = DateTime.UtcNow + maxWait;
            int stable = 0;
            int lastCount = -1;

            while (DateTime.UtcNow < end)
            {
                try
                {
                    try { js.ExecuteScript("window.scrollTo(0, 60);"); } catch { }
                    System.Threading.Thread.Sleep(120);
                    try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                }
                catch { }

                int visibleTopAds = 0;
                try
                {
                    var script = @"
                        const sels = [
                          'header .banner', 'header [id*=\\'banner\\']', 'header [class*=\\'banner\\']',
                          '#banner_top', '.banner-top', '.top-banner', '.leaderboard', '.top-ads', '.banner-leaderboard',
                          '[id^=\\'div-gpt-ad\\']', '[id*=\\'gpt-ad\\']', '.gpt-ad', '.dfp-ad', '.ad-slot', '.ad-container',
                          '[id*=\\'ads\\']', '[class*=\\'ad-\\']', '[class*=\\'-ad\\']', '.ads', '.advertisement'
                        ].join(',');
                        const minW = 120, minH = 30, maxY = 900;
                        const y = window.scrollY || window.pageYOffset || 0;
                        let count = 0;
                        const list = Array.from(document.querySelectorAll(sels));
                        for (const el of list) {
                          try {
                            const r = el.getBoundingClientRect();
                            const w = Math.round(r.width), h = Math.round(r.height);
                            const top = Math.round(r.top + y);
                            if (w >= minW && h >= minH && top < maxY) count++;
                          } catch {}
                        }
                        const ifr = Array.from(document.querySelectorAll('iframe,frame'));
                        for (const el of ifr) {
                          try {
                            const r = el.getBoundingClientRect();
                            const w = Math.round(r.width), h = Math.round(r.height);
                            const top = Math.round(r.top + y);
                            if (w >= minW && h >= minH && top < maxY) count++;
                          } catch {}
                        }
                        return count;
                    ";
                    visibleTopAds = Convert.ToInt32(js.ExecuteScript(script) ?? 0);
                }
                catch { }

                if (visibleTopAds <= 0)
                {
                    System.Threading.Thread.Sleep(220);
                    continue;
                }

                if (visibleTopAds == lastCount) stable++; else { stable = 0; lastCount = visibleTopAds; }
                if (stable >= 2) break;

                System.Threading.Thread.Sleep(180);
            }
        }

        /// <summary>
        /// Tự động tìm đường dẫn Chrome/Chromium binary trên hệ thống
        /// 
        /// Cách thức hoạt động:
        /// 1. Kiểm tra hệ điều hành (macOS, Linux, Windows)
        /// 2. Với mỗi OS, thử các đường dẫn phổ biến:
        /// 
        /// macOS:
        ///    - /Applications/Google Chrome.app/Contents/MacOS/Google Chrome
        ///    - /Applications/Chromium.app/Contents/MacOS/Chromium
        /// 
        /// Linux:
        ///    - /usr/bin/google-chrome
        ///    - /usr/bin/google-chrome-stable
        ///    - /usr/bin/chromium
        ///    - /usr/bin/chromium-browser
        /// 
        /// Windows:
        ///    - ProgramFiles\Google\Chrome\Application\chrome.exe
        ///    - ProgramFiles(x86)\Google\Chrome\Application\chrome.exe
        ///    - LocalApplicationData\Google\Chrome\Application\chrome.exe
        ///    - ProgramFiles\Chromium\Application\chrome.exe
        ///    - ProgramFiles(x86)\Chromium\Application\chrome.exe
        ///    - Registry: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe
        ///    - Registry: HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe
        /// 
        /// 3. Trả về đường dẫn đầu tiên tìm thấy file tồn tại
        /// 4. Nếu không tìm thấy => trả về chuỗi rỗng
        /// 
        /// Mục đích: Tự động phát hiện Chrome để không cần cấu hình thủ công
        /// 
        /// Trả về: Đường dẫn đầy đủ đến Chrome binary, hoặc chuỗi rỗng nếu không tìm thấy
        /// </summary>
        private static string FindChromeBinaryPath()
        {
            try
            {
                // macOS paths
                if (OperatingSystem.IsMacOS())
                {
                    var macCandidates = new[]
                    {
                        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                        "/Applications/Chromium.app/Contents/MacOS/Chromium",
                        "/Users/lecuong/Desktop/Google Chrome.app/Contents/MacOS/Google Chrome"
                    };
                    foreach (var p in macCandidates)
                    {
                        if (File.Exists(p)) return p;
                    }
                    
                    // Tìm Chrome for Testing trong cache (Selenium tự download)
                    try
                    {
                        var cacheBase = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".cache", "selenium", "chrome");
                        if (Directory.Exists(cacheBase))
                        {
                            // Tìm tất cả các version trong cache
                            var platformDirs = Directory.GetDirectories(cacheBase);
                            foreach (var platformDir in platformDirs)
                            {
                                var versionDirs = Directory.GetDirectories(platformDir);
                                foreach (var versionDir in versionDirs)
                                {
                                    var chromePath = Path.Combine(versionDir, "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing");
                                    if (File.Exists(chromePath))
                                    {
                                        return chromePath;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                    
                    return string.Empty;
                }

                // Linux paths
                if (OperatingSystem.IsLinux())
                {
                    var linuxCandidates = new[]
                    {
                        "/usr/bin/google-chrome",
                        "/usr/bin/google-chrome-stable",
                        "/usr/bin/chromium",
                        "/usr/bin/chromium-browser"
                    };
                    foreach (var p in linuxCandidates)
                    {
                        if (File.Exists(p)) return p;
                    }
                    return string.Empty;
                }

                // Windows paths
                if (OperatingSystem.IsWindows())
                {
                    var windowsCandidates = new[]
                    {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Google\\Chrome\\Application\\chrome.exe"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Google\\Chrome\\Application\\chrome.exe"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google\\Chrome\\Application\\chrome.exe"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Chromium\\Application\\chrome.exe"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Chromium\\Application\\chrome.exe")
                    };
                    foreach (var p in windowsCandidates)
                    {
                        if (File.Exists(p)) return p;
                    }

                    // Registry lookup for Windows
                    string[] hives = { "HKEY_LOCAL_MACHINE", "HKEY_CURRENT_USER" };
                    foreach (var hive in hives)
                    {
                        var path = Registry.GetValue($"{hive}\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe", "", null) as string;
                        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) return path;
                    }
                }
            }
            catch { }
            return string.Empty;
        }

        /// <summary>
        /// Thử điều hướng tới URL với nhiều phương pháp và retry
        /// 
        /// Cách thức hoạt động:
        /// 1. Có tối đa 2 lần thử (attempt = 1, 2)
        /// 2. Lần thử 1: Dùng driver.Navigate().GoToUrl(url)
        ///    - Đợi document readyState = "complete" hoặc "interactive" trong 20 giây
        ///    - Nếu thành công => return true
        ///    - Nếu fail => lưu lỗi và tiếp tục
        /// 3. Lần thử 2: Dùng JavaScript redirect (window.location.href = url)
        ///    - Đôi khi phương pháp này thành công khi GoToUrl() bị block
        ///    - Đợi readyState trong 25 giây
        ///    - Nếu thành công => return true
        /// 4. Nếu cả 2 lần đều fail:
        ///    - Kiểm tra URL hiện tại của driver
        ///    - Nếu có URL => thêm vào failureReason
        ///    - Ghi log exception cuối cùng
        ///    - Return false
        /// 
        /// Xử lý exception:
        /// - WebDriverTimeoutException: Timeout khi điều hướng
        /// - WebDriverException: Lỗi WebDriver
        /// - Exception khác: Lỗi chung
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - url: URL cần điều hướng tới
        /// - timeout: Thời gian timeout tổng
        /// - failureReason: Lý do thất bại (output parameter)
        /// 
        /// Trả về: true nếu điều hướng thành công, false nếu thất bại
        /// </summary>
        private static bool TryNavigate(IWebDriver driver, string url, TimeSpan timeout, out string failureReason)
        {
            var deadline = DateTime.UtcNow + timeout;
            failureReason = string.Empty;
            Exception? lastEx = null;
            for (int attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    // Attempt standard navigation
                    driver.Navigate().GoToUrl(url);
                    if (WaitForReadyState(driver, TimeSpan.FromSeconds(20))) return true;

                    // If readyState not achieved, treat as failure and retry once with JS redirect
                    failureReason = "Document not ready after navigation";
                }
                catch (WebDriverTimeoutException tex)
                {
                    lastEx = tex;
                    failureReason = tex.Message;
                }
                catch (WebDriverException wex)
                {
                    lastEx = wex;
                    failureReason = wex.Message;
                }
                catch (Exception ex)
                {
                    lastEx = ex;
                    failureReason = ex.Message;
                }

                if (DateTime.UtcNow > deadline) break;

                try
                {
                    // Retry via JS location assignment which is sometimes more reliable behind blockers
                    ((IJavaScriptExecutor)driver).ExecuteScript("window.location.href = arguments[0];", url);
                }
                catch (Exception jsex)
                {
                    lastEx = jsex;
                    failureReason = $"JS redirect failed: {jsex.Message}";
                }

                if (WaitForReadyState(driver, TimeSpan.FromSeconds(25))) return true;
            }

            if (!string.IsNullOrWhiteSpace(driver?.Url))
            {
                failureReason = string.IsNullOrWhiteSpace(failureReason)
                    ? $"Arrived at unexpected URL: {driver.Url}"
                    : failureReason + $" | current URL: {driver.Url}";
            }
            if (lastEx != null)
            {
                ErrorWriter.WriteLog(LogPath, "TryNavigateException", lastEx.ToString());
            }
            return false;
        }

        /// <summary>
        /// Đợi trang web đạt trạng thái ready (đã load xong)
        /// 
        /// Cách thức hoạt động:
        /// 1. Tạo WebDriverWait với thời gian chờ = wait, poll interval = 250ms
        /// 2. Liên tục kiểm tra document.readyState bằng JavaScript:
        ///    - "complete": Trang đã load hoàn toàn
        ///    - "interactive": DOM đã sẵn sàng, nhưng có thể còn resource đang load
        /// 3. Nếu readyState = "complete" hoặc "interactive" => return true
        /// 4. Nếu quá timeout => return false
        /// 5. Nếu có lỗi khi chạy JavaScript => return false
        /// 
        /// Lưu ý: Đây là cách tiêu chuẩn để đợi trang load trước khi thao tác
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - wait: Thời gian tối đa để đợi
        /// 
        /// Trả về: true nếu trang đã ready, false nếu timeout hoặc lỗi
        /// </summary>
        private static bool WaitForReadyState(IWebDriver driver, TimeSpan wait)
        {
            try
            {
                var waitUntil = new WebDriverWait(new SystemClock(), driver, wait, TimeSpan.FromMilliseconds(250));
                return waitUntil.Until(d =>
                {
                    try
                    {
                        var state = ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState")?.ToString();
                        return state == "complete" || state == "interactive";
                    }
                    catch
                    {
                        return false;
                    }
                });
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Trả về danh sách CSS selector chung để tìm quảng cáo trên hầu hết website
        /// 
        /// Cách thức hoạt động:
        /// Trả về mảng các selector phổ biến, bao gồm:
        /// 1. Selector theo id/class chứa từ khóa:
        ///    - ad, ads, advert, banner (ví dụ: [id*='ad'], [class*='ad-'])
        /// 2. Selector của các platform quảng cáo phổ biến:
        ///    - Google DFP/GPT: .gpt-ad, .gpt-unit, .gpt-slot, .dfp-ad, [id^='div-gpt-ad']
        ///    - Google AdSense: .adsbygoogle, .google-auto-placed
        ///    - Ad container: .ad-slot, .ad-container, .ad-wrapper, .ad__container
        /// 3. Selector quảng cáo tự nhiên/sponsored:
        ///    - .sponsor, .sponsored, .promoted, .native-ad, .in-content-ad
        /// 4. Selector theo vị trí:
        ///    - .header-ad, .footer-ad, .sidebar-ad, .sticky-ad, .floating-ad
        /// 5. Selector theo data attribute:
        ///    - [data-ad], [data-ad-slot], [data-ad-unit], [data-google-query-id]
        /// 
        /// Mục đích: Tập hợp selector phổ biến nhất để tìm quảng cáo trên đa số website
        /// 
        /// Trả về: Mảng string chứa các CSS selector
        /// </summary>
        private static string[] GetCommonAdSelectors()
        {
            return new[]
            {
                // Theo id/class chứa từ khóa ad/ads/advert/banner
                "[id*='ad'], [id*='ads'], [id*='advert'], [id*='banner']",
                "[class*='ad '], [class*=' ad'], [class*='ad-'], [class*='-ad'], [class*='ads'], [class*='advert'], [class*='banner']",

                // Các lớp phổ biến của DFP/GPT/AdSense và các CMS
                ".gpt-ad, .gpt-unit, .gpt-slot, .dfp-ad, .dfp-slot, .ad-slot, .ad-container, .ad-wrapper, .ad__container, .ad__slot, .adsbygoogle, .google-auto-placed",

                // Biểu thị tài trợ/quảng cáo tự nhiên
                ".sponsor, .sponsored, .promoted, .native-ad, .in-content-ad, .article-ad, .article-advertisement",

                // Vị trí thường gặp
                ".header-ad, .footer-ad, .sidebar-ad, .sticky-ad, .floating-ad, #footer-ad, .footer-banner, .footer-advertisement",

                // Theo data-* phổ biến
                "[data-ad], [data-ad-slot], [data-ad-unit], [data-google-query-id], [data-ez-name]"
            };
        }

        /// <summary>
        /// Cuộn trang xuống cuối để kích hoạt lazy-loading và đảm bảo nội dung đã load hết
        /// 
        /// Cách thức hoạt động:
        /// 1. Lấy chiều cao ban đầu của trang (document.scrollHeight)
        /// 2. Trong khoảng thời gian maxWait, lặp lại:
        ///    - Cuộn xuống cuối trang (scrollTo(0, document.body.scrollHeight))
        ///    - Đợi 500ms để nội dung lazy-load render
        ///    - Lấy chiều cao mới của trang
        ///    - So sánh với chiều cao cũ:
        ///      * Nếu bằng nhau => stableRounds++
        ///      * Nếu khác nhau => reset stableRounds = 0, cập nhật lastHeight
        ///    - Nếu stableRounds >= 3 (chiều cao không đổi trong 3 lần liên tiếp) => thoát
        /// 3. Sau khi thoát: Đợi thêm readyState = "complete" trong 2 giây để chắc chắn
        /// 
        /// Mục đích: 
        /// - Kích hoạt lazy-loading: Nhiều website load nội dung khi scroll xuống
        /// - Đảm bảo quảng cáo ở cuối trang đã được load trước khi chụp ảnh
        /// - Tối ưu: Dừng sớm khi không còn nội dung mới load
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - maxWait: Thời gian tối đa để cuộn và đợi
        /// </summary>
        private static void ScrollToBottomAndEnsureLazyContent(IWebDriver driver, TimeSpan maxWait)
        {
            var js = (IJavaScriptExecutor)driver;
            long lastHeight = 0;
            int stableRounds = 0;
            var deadline = DateTime.UtcNow + maxWait;

            try
            {
                lastHeight = Convert.ToInt64(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
            }
            catch { }

            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
                }
                catch { }

                Thread.Sleep(500);

                long newHeight = lastHeight;
                try
                {
                    newHeight = Convert.ToInt64(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
                }
                catch { }

                if (newHeight <= lastHeight)
                {
                    stableRounds++;
                }
                else
                {
                    stableRounds = 0;
                    lastHeight = newHeight;
                }

                if (stableRounds >= 3) break;
            }

            // Một lần chờ ngắn để JS còn lại hoàn tất
            WaitForReadyState(driver, TimeSpan.FromSeconds(2));
        }

        /// <summary>
        /// Thăm dò xem các phần tử quảng cáo đã xuất hiện trên trang chưa (hàm phụ trợ)
        /// 
        /// Cách thức hoạt động:
        /// 1. Trong khoảng thời gian wait, liên tục:
        /// 2. Chạy JavaScript để đếm số lượng element quảng cáo:
        ///    - Selector: div.banner, section.banner, #banner, [id*='banner'], 
        ///                [class*='banner'], div.ads, .ads, [id*='ads'], 
        ///                [class*='ad-'], [class*='ads-'], div.advertisement, 
        ///                .advertisement, [class*='advert']
        ///    - Trả về tổng số element tìm thấy
        /// 3. So sánh với lần kiểm tra trước:
        ///    - Nếu bằng nhau => stable++
        ///    - Nếu khác nhau => reset stable = 0, cập nhật lastCount
        /// 4. Nếu stable >= 2 (số lượng ổn định trong 2 lần liên tiếp) => thoát
        /// 5. Đợi 300ms trước khi kiểm tra lại
        /// 
        /// Mục đích: Thăm dò nhanh xem quảng cáo đã render chưa, không đợi quá lâu
        /// Khác với WaitForAdsLoaded: Hàm này chỉ thăm dò, không chờ quá kỹ
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - wait: Thời gian tối đa để thăm dò
        /// </summary>
        private static void TryProbeAdCandidates(IWebDriver driver, TimeSpan wait)
        {
            var js = (IJavaScriptExecutor)driver;
            var end = DateTime.UtcNow + wait;
            int lastCount = -1;
            int stable = 0;
            while (DateTime.UtcNow < end)
            {
                int count = 0;
                try
                {
                    // Các selector khả dĩ chung cho banner/ads
                    var script = "return document.querySelectorAll(\"div.banner, section.banner, #banner, [id*='banner'], [class*='banner'], div.ads, .ads, [id*='ads'], [class*='ad-'], [class*='ads-'], div.advertisement, .advertisement, [class*='advert']\").length;";
                    count = Convert.ToInt32(js.ExecuteScript(script));
                }
                catch { }

                if (count == lastCount)
                {
                    stable++;
                }
                else
                {
                    stable = 0;
                    lastCount = count;
                }

                if (stable >= 2) break;
                Thread.Sleep(300);
            }
        }
    }

    
}