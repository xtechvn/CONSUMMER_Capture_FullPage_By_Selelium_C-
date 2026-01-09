// Create By: cuonglv 2 2
// git add .
// git commit -m "update code mới"
// git push origin main

// git add ConsummerScreenPageBot/Program.cs
// git commit -m "Update Program.cs"
// git push origin main

// keo de code
// git fetch --all
//git reset --hard origin/main

using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Configuration;
using System.Linq;
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
using System.Diagnostics;

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
         public static string task_bot = ConfigurationManager.AppSettings["task"] ?? "";
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

             
                // Đảm bảo thư mục log tồn tại trước khi yêu cầu ChromeDriver ghi log
                try
                {
                    if (!string.IsNullOrWhiteSpace(LogPath))
                    {
                        Directory.CreateDirectory(LogPath);
                    }
                }
                catch (Exception logDirEx)
                {
                    Console.WriteLine($"[Chrome] Cannot prepare log directory '{LogPath}': {logDirEx.Message}");
                    throw;
                }

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

                                    // Link web can chup: screen_banner: chụp banner , link_click_banner: chụp trang đích sau khi click banner
                                    // Fix null pointer if property does not exist 277
                                    if (task_bot == "screen_banner")
                                    {
                                        siteUrl = jobj["link_web"] != null ? jobj["link_web"].ToObject<string>() : "";
                                    }
                                    else
                                    {
                                        siteUrl = jobj["link_click_banner"] != null ? jobj["link_click_banner"].ToObject<string>() : "";
                                    }

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
        /// Chụp ảnh trang web bằng cách chia thành nhiều segment (phần) và chụp từng phần bằng cách scroll đến từng vị trí
        /// 
        /// Cách thức hoạt động:
        /// 1. Scroll xuống cuối trang để kích hoạt lazy-load và đợi quảng cáo load hoàn toàn
        /// 2. Scroll lên đầu trang để reset vị trí
        /// 3. Tính chiều cao mỗi segment: totalHeight / segmentCount
        /// 4. Với mỗi segment i:
        ///    - Scroll đến vị trí Y = i * sliceHeight (đặt ở giữa viewport)
        ///    - Đợi một chút để đảm bảo content đã render
        ///    - Chụp viewport tại vị trí đó
        ///    - Lưu segment với tên file: split{i+1}_YYYYMMDD_HHmmss_ffffff_guid.jpg
        ///    - Nén ảnh với chất lượng jpegQuality
        ///    - Gửi ảnh lên queue analyze để AI xử lý
        /// 5. Scroll lại lên đầu trang sau khi xong
        /// 
        /// Ưu điểm: Mỗi segment là ảnh chụp viewport riêng biệt khi scroll, đảm bảo không bỏ sót nội dung
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
                
                // Bước 1: Scroll xuống cuối trang để kích hoạt lazy-load
                Console.WriteLine("[Segment] Scroll xuống cuối trang để kích hoạt lazy-load...");
                ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(8));
                
                // Bước 2: Đợi quảng cáo load hoàn toàn
                WaitForAdsLoaded(driver, TimeSpan.FromSeconds(6));
                
                // Bước 3: Lấy chiều cao tổng của trang
                int totalHeight = 3000;
                int viewportHeight = 1080;
                try
                {
                    totalHeight = Convert.ToInt32(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight, window.innerHeight || 0);"));
                    viewportHeight = Convert.ToInt32(js.ExecuteScript("return window.innerHeight || document.documentElement.clientHeight || 1080;"));
                }
                catch { }
                if (totalHeight <= 0) totalHeight = 3000;
                totalHeight = Math.Min(totalHeight, 30000);
                if (viewportHeight <= 0) viewportHeight = 1080;

                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                // Bước 4: Tính chiều cao mỗi segment và scroll đến từng vị trí để chụp
                int sliceHeight = (int)Math.Ceiling((double)totalHeight / segmentCount);
                Console.WriteLine($"[Segment] Total height: {totalHeight}px, Viewport height: {viewportHeight}px, Segment height: {sliceHeight}px");

                // Scroll lên đầu trang trước khi bắt đầu
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                Thread.Sleep(500);

                for (int i = 0; i < segmentCount; i++)
                {
                    // Tính vị trí Y để scroll đến (đặt vị trí segment ở giữa viewport)
                    int targetY = i * sliceHeight;
                    // Điều chỉnh để segment nằm ở giữa viewport
                    int scrollY = Math.Max(0, targetY - (viewportHeight / 2));
                    
                    Console.WriteLine($"[Segment] Segment {i + 1}/{segmentCount}: Scroll to Y={scrollY} (target segment at Y={targetY})");
                    
                    // Scroll đến vị trí
                    try
                    {
                        js.ExecuteScript($"window.scrollTo(0, {scrollY});");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Segment] Error scrolling to {scrollY}: {ex.Message}");
                    }
                    
                    // Đợi để đảm bảo content đã render
                    Thread.Sleep(800);
                    
                    // Chụp viewport tại vị trí này
                    byte[] segmentBytes = Array.Empty<byte>();
                    try
                    {
                        var shot = ((ITakesScreenshot)driver).GetScreenshot();
                        segmentBytes = shot.AsByteArray;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Segment] Error capturing segment {i + 1}: {ex.Message}");
                        continue;
                    }
                    
                    if (segmentBytes == null || segmentBytes.Length == 0)
                    {
                        Console.WriteLine($"[Segment] WARNING: Segment {i + 1} screenshot is empty");
                        continue;
                    }
                    
                    // Load ảnh và lưu
                    try
                    {
                        using (var ms = new MemoryStream(segmentBytes))
                        using (var img = Image.Load<Rgba32>(ms))
                        {
                            var fileName = $"split{i+1}_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.jpg";
                            var savePath = Path.Combine(shotsDir, fileName);
                            
                            // Nén ảnh
                            AdCapture.SaveImageCompressed(img, savePath, 1.0, jpegQuality);
                            
                            var compressedBytes = File.ReadAllBytes(savePath);
                            Console.WriteLine($"[Segment] Saved segment {i + 1}/{segmentCount}: {Path.GetFileName(savePath)} ({compressedBytes.Length} bytes, {img.Width}x{img.Height}px)");
                            
                            // Gửi lên queue
                            TryPublishAnalyze(compressedBytes, img.Width, img.Height);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Segment] Error processing segment {i + 1}: {ex.Message}");
                        ErrorWriter.WriteLog(LogPath, "SegmentScreenshots.Process", $"Segment {i + 1} - {hostLabel} => {ex}");
                    }
                }
                
                // Scroll lại lên đầu trang sau khi xong
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                
                Console.WriteLine($"[Segment] Done host={hostLabel}, captured {segmentCount} segments");
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
        /// <summary>
        /// Normalize URL để so sánh và tránh lặp:
        /// - Loại bỏ query params (utm_source, tracking params, ...)
        /// - Loại bỏ fragment (#...)
        /// - Normalize trailing slash
        /// - Lowercase host
        /// </summary>
        private static string NormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return url ?? "";
            }
            
            try
            {
                var uri = new Uri(url);
                // Loại bỏ query params và fragment, chỉ giữ scheme://host/path
                string path = uri.AbsolutePath;
                // Normalize trailing slash: loại bỏ trailing slash (trừ root path)
                if (path.EndsWith("/") && path.Length > 1)
                {
                    path = path.TrimEnd('/');
                }
                
                string normalized = $"{uri.Scheme.ToLowerInvariant()}://{uri.Host.ToLowerInvariant()}{path}";
                return normalized;
            }
            catch
            {
                // Nếu không parse được URL, trả về nguyên bản đã trim và lowercase
                return url.Trim().ToLowerInvariant();
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
                // Kiểm tra nếu nhiem vu cua Bot la chup page theo link  thì bỏ qua hàm này
                if (task_bot == "screen_link_page")
                {
                    Console.WriteLine("[AdLink] RabbitQueueScreenLink có chứa 'screen_link_page', bỏ qua ProcessIframesAndPushToQueue.");
                    return;
                }
                
                Console.WriteLine($"[AdLink] Bắt đầu xử lý iframe và element quảng cáo cho host={hostLabel}");
                
                // Lưu window handle gốc
                string originalWindowHandle = driver.CurrentWindowHandle;
                string originalUrl = driver.Url;
                
                // Lấy domain chính của trang hiện tại để loại bỏ link cùng domain
                // Ví dụ: tuoitre.vn -> lấy "tuoitre", vnexpress.net -> lấy "vnexpress"
                string currentDomainKey = string.Empty;
                try
                {
                    var uri = new Uri(originalUrl);
                    string host = uri.Host.ToLowerInvariant();
                    // Loại bỏ www. nếu có
                    if (host.StartsWith("www."))
                    {
                        host = host.Substring(4);
                    }
                    // Lấy phần domain chính (trước dấu chấm đầu tiên)
                    int firstDotIndex = host.IndexOf('.');
                    if (firstDotIndex > 0)
                    {
                        currentDomainKey = host.Substring(0, firstDotIndex);
                    }
                    else
                    {
                        currentDomainKey = host;
                    }
                    Console.WriteLine($"[AdLink] Domain key hiện tại: {currentDomainKey} (từ {host}), sẽ loại bỏ các link chứa domain này");
                }
                catch
                {
                    Console.WriteLine($"[AdLink] Không thể lấy domain từ URL: {originalUrl}");
                }
                
                // Đảm bảo chỉ giữ lại tab gốc trước khi bắt đầu xử lý
                CloseAllNonOriginalTabs(driver, ref originalWindowHandle);

                // Hàm helper để kiểm tra link có chứa domain key hiện tại không
                bool IsExternalLink(string link)
                {
                    if (string.IsNullOrWhiteSpace(link) || string.IsNullOrWhiteSpace(currentDomainKey))
                        return true; // Nếu không có domain key để so sánh, cho phép tất cả
                    
                    string linkLower = link.ToLowerInvariant();
                    // Nếu IndexOf = -1 nghĩa là không chứa domain key, đây là link external
                    return linkLower.IndexOf(currentDomainKey, StringComparison.OrdinalIgnoreCase) == -1;
                }
                
                // Lấy job params từ RabbitQueue gốc để merge vào payload
                var jobParamsSnapshot = lastJobParams; // tránh race condition
                
                int processedCount = 0;
                int handledBannerOrdinal = 0;
                var processedFingerprints = new HashSet<string>(StringComparer.Ordinal);
                var processedUrls = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // Track URL đã xử lý để tránh lặp
                
                while (true)
                {
                    var bannerElements = CollectBannerElements(driver);
                    if (bannerElements.Count == 0)
                    {
                        if (handledBannerOrdinal == 0)
                        {
                            Console.WriteLine("[AdLink] Không tìm thấy banner nào, kết thúc xử lý");
                        }
                        else
                        {
                            Console.WriteLine($"[AdLink] Không còn banner mới sau khi đã xử lý {handledBannerOrdinal} banner(s)");
                        }
                        break;
                    }

                    var bannerElement = GetNextBannerElement(driver, bannerElements, processedFingerprints, out var bannerFingerprint, out var bannerPositionIndex);
                    if (bannerElement == null)
                    {
                        Console.WriteLine("[AdLink] Không còn banner mới để xử lý trong lần quét này");
                        break;
                    }

                    handledBannerOrdinal++;
                    Console.WriteLine($"[AdLink] Đang xử lý banner {handledBannerOrdinal}. Vị trí hiện tại {bannerPositionIndex + 1}/{bannerElements.Count}");

                    bool shouldMarkFingerprint = true;
                    try
                    {
                        // Đảm bảo đang ở tab gốc
                        try
                        {
                            driver.SwitchTo().Window(originalWindowHandle);
                        }
                        catch
                        {
                            if (driver.WindowHandles.Contains(originalWindowHandle))
                            {
                                driver.SwitchTo().Window(originalWindowHandle);
                            }
                            else
                            {
                                Console.WriteLine($"[AdLink] Tab gốc không còn tồn tại, bỏ qua banner {handledBannerOrdinal}");
                                shouldMarkFingerprint = false;
                                continue;
                            }
                        }

                        // Dọn dẹp các tab thừa trước khi click banner mới
                        CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                        
                        // Lưu số lượng window hiện tại
                        var windowsBefore = driver.WindowHandles.Count;
                        string? targetUrl = null;
                        bool shouldOpenUrl = false;
                        
                        try
                        {
                            // Scroll đến element để đảm bảo nó visible
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", bannerElement);
                            System.Threading.Thread.Sleep(500);
                            
                            // Lấy tọa độ của element (center point) bằng JavaScript
                            var js = (IJavaScriptExecutor)driver;
                            var rect = js.ExecuteScript(@"
                                var el = arguments[0];
                                var rect = el.getBoundingClientRect();
                                var centerX = Math.round(rect.left + rect.width / 2);
                                var centerY = Math.round(rect.top + rect.height / 2);
                                return {
                                    clientX: centerX,
                                    clientY: centerY,
                                    width: Math.round(rect.width),
                                    height: Math.round(rect.height)
                                };
                            ", bannerElement) as System.Collections.Generic.Dictionary<string, object>;
                            
                            if (rect != null && rect.ContainsKey("clientX") && rect.ContainsKey("clientY"))
                            {
                                int centerX = Convert.ToInt32(rect["clientX"]);
                                int centerY = Convert.ToInt32(rect["clientY"]);
                                int width = Convert.ToInt32(rect["width"]);
                                int height = Convert.ToInt32(rect["height"]);
                                
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Tọa độ center (viewport): ({centerX}, {centerY}), kích thước: {width}x{height}");
                                
                                bool isIframe = bannerElement.TagName.ToLower() == "iframe";
                                
                                bool iframeClickedByViewport = false;
                                if (isIframe)
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Thử click iframe tại ({centerX}, {centerY})");
                                    iframeClickedByViewport = TryClickIframeElement(driver, bannerElement, centerX, centerY);
                                }
                                else
                                {
                                    try
                                    {
                                        var links = bannerElement.FindElements(By.TagName("a"));
                                        foreach (var link in links)
                                        {
                                            try
                                            {
                                                string? href = link.GetAttribute("href");
                                                if (!string.IsNullOrWhiteSpace(href) && IsValidHttpLink(href) && IsExternalLink(href))
                                                {
                                                    targetUrl = href;
                                                    shouldOpenUrl = true;
                                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Tìm thấy link: {href.Substring(0, Math.Min(80, href.Length))}...");
                                                    break;
                                                }
                                            }
                                            catch { }
                                        }
                                        
                                        if (string.IsNullOrWhiteSpace(targetUrl))
                                        {
                                            string? dataHref = bannerElement.GetAttribute("data-href");
                                            if (!string.IsNullOrWhiteSpace(dataHref) && IsValidHttpLink(dataHref) && IsExternalLink(dataHref))
                                            {
                                                targetUrl = dataHref;
                                                shouldOpenUrl = true;
                                            }
                                            else
                                            {
                                                string? dataUrl = bannerElement.GetAttribute("data-url");
                                                if (!string.IsNullOrWhiteSpace(dataUrl) && IsValidHttpLink(dataUrl) && IsExternalLink(dataUrl))
                                                {
                                                    targetUrl = dataUrl;
                                                    shouldOpenUrl = true;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[AdLink] Lỗi khi tìm link trong banner {handledBannerOrdinal}: {ex.Message}");
                                    }
                                }
                                
                                if (string.IsNullOrWhiteSpace(targetUrl))
                                {
                                    try
                                    {
                                        if (isIframe && !iframeClickedByViewport)
                                        {
                                            try
                                            {
                                                driver.SwitchTo().Frame(bannerElement);
                                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Đã switch vào iframe");
                                                System.Threading.Thread.Sleep(1000);
                                                
                                                var linksInFrame = driver.FindElements(By.TagName("a"));
                                                bool foundLink = false;
                                                
                                                foreach (var link in linksInFrame)
                                                {
                                                    try
                                                    {
                                                        string? href = link.GetAttribute("href");
                                                        if (!string.IsNullOrWhiteSpace(href) && IsValidHttpLink(href) && IsExternalLink(href))
                                                        {
                                                            targetUrl = href;
                                                            shouldOpenUrl = true;
                                                            foundLink = true;
                                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Tìm thấy link: {href.Substring(0, Math.Min(80, href.Length))}...");
                                                            break;
                                                        }
                                                    }
                                                    catch { }
                                                }
                                                
                                                if (!foundLink)
                                                {
                                                    try
                                                    {
                                                        var frameSize = js.ExecuteScript(@"
                                                            return {
                                                                width: document.body.scrollWidth || document.documentElement.clientWidth,
                                                                height: document.body.scrollHeight || document.documentElement.clientHeight
                                                            };
                                                        ") as System.Collections.Generic.Dictionary<string, object>;
                                                        
                                                        if (frameSize != null)
                                                        {
                                                            int frameWidth = Convert.ToInt32(frameSize["width"]);
                                                            int frameHeight = Convert.ToInt32(frameSize["height"]);
                                                            int frameCenterX = frameWidth / 2;
                                                            int frameCenterY = frameHeight / 2;
                                                            
                                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Click vào center của iframe content ({frameCenterX}, {frameCenterY})");
                                                            
                                                            var clickResult = js.ExecuteScript(@"
                                                                var x = arguments[0];
                                                                var y = arguments[1];
                                                                var element = document.elementFromPoint(x, y);
                                                                if (!element) return {clicked: false};
                                                                
                                                                var clickable = element;
                                                                while (clickable && clickable !== document.body) {
                                                                    var tag = clickable.tagName.toLowerCase();
                                                                    var hasHref = clickable.href || clickable.getAttribute('href');
                                                                    
                                                                    if (tag === 'a' || hasHref) {
                                                                        var href = clickable.href || clickable.getAttribute('href') || '';
                                                                        if (href) {
                                                                            window.open(href, '_blank');
                                                                            return {clicked: true, href: href};
                                                                        }
                                                                    }
                                                                    clickable = clickable.parentElement;
                                                                }
                                                                
                                                                var event = new MouseEvent('click', {
                                                                    view: window,
                                                                    bubbles: true,
                                                                    cancelable: true,
                                                                    clientX: x,
                                                                    clientY: y,
                                                                    button: 0
                                                                });
                                                                element.dispatchEvent(event);
                                                                return {clicked: true, href: ''};
                                                            ", frameCenterX, frameCenterY) as System.Collections.Generic.Dictionary<string, object>;
                                                            
                                                            if (clickResult != null && clickResult.ContainsKey("clicked"))
                                                            {
                                                                bool clicked = Convert.ToBoolean(clickResult["clicked"]);
                                                                string href = clickResult.ContainsKey("href") ? clickResult["href"].ToString() ?? "" : "";
                                                                
                                                                if (clicked && !string.IsNullOrWhiteSpace(href) && IsValidHttpLink(href) && IsExternalLink(href))
                                                                {
                                                                    targetUrl = href;
                                                                    shouldOpenUrl = true;
                                                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Tìm thấy href từ click: {href.Substring(0, Math.Min(80, href.Length))}...");
                                                                }
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex2)
                                                    {
                                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Lỗi khi click trong iframe: {ex2.Message}");
                                                    }
                                                }
                                                
                                                driver.SwitchTo().DefaultContent();
                                            }
                                            catch (Exception exActions)
                                            {
                                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} (iframe) -> Lỗi khi xử lý iframe: {exActions.Message}");
                                                try { driver.SwitchTo().DefaultContent(); } catch { }
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Thử click trực tiếp vào banner");
                                            
                                            bool clicked = false;
                                            
                                            try
                                            {
                                                var actions = new Actions(driver);
                                                actions.MoveToElement(bannerElement).Click().Perform();
                                                clicked = true;
                                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Click thành công bằng Actions");
                                            }
                                            catch (Exception)
                                            {
                                                try
                                                {
                                                    js.ExecuteScript("arguments[0].click();", bannerElement);
                                                    clicked = true;
                                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Click thành công bằng JavaScript");
                                                }
                                                catch (Exception)
                                                {
                                                    try
                                                    {
                                                        bannerElement.Click();
                                                        clicked = true;
                                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Click thành công trực tiếp");
                                                    }
                                                    catch (Exception)
                                                    {
                                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Tất cả cách click đều fail");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[AdLink] Lỗi khi click vào tọa độ banner {handledBannerOrdinal}: {ex.Message}");
                                        try
                                        {
                                            var actions = new Actions(driver);
                                            actions.MoveToElement(bannerElement).Click().Perform();
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Fallback: Click bằng Actions");
                                        }
                                        catch (Exception ex2)
                                        {
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Fallback Actions cũng fail: {ex2.Message}");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Không thể lấy tọa độ, thử click trực tiếp");
                                try
                                {
                                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", bannerElement);
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[AdLink] Lỗi khi xử lý banner {handledBannerOrdinal}: {ex.Message}");
                            continue;
                        }
                        
                        if (shouldOpenUrl && !string.IsNullOrWhiteSpace(targetUrl))
                        {
                            try
                            {
                                ((IJavaScriptExecutor)driver).ExecuteScript("window.open(arguments[0], '_blank');", targetUrl);
                                System.Threading.Thread.Sleep(800);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[AdLink] Lỗi khi mở URL cho banner {handledBannerOrdinal}: {ex.Message}");
                                continue;
                            }
                        }
                        
                        string? newWindowHandle;
                        bool hasNewTab = TryGetNewWindowHandle(driver, originalWindowHandle, windowsBefore, TimeSpan.FromSeconds(6), out newWindowHandle);
                        TimeSpan pageLoadTimeout = TimeSpan.FromSeconds(6);

                        if (hasNewTab && !string.IsNullOrWhiteSpace(newWindowHandle))
                        {
                            try
                            {
                                driver.SwitchTo().Window(newWindowHandle);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[AdLink] Không thể switch sang tab mới: {ex.Message}");
                                driver.SwitchTo().Window(originalWindowHandle);
                                continue;
                            }
                            
                            if (!WaitForPageLoad(driver, pageLoadTimeout))
                            {
                                Console.WriteLine($"[AdLink] Tab mới của banner {handledBannerOrdinal} load quá {pageLoadTimeout.TotalSeconds}s, đóng tab và bỏ qua");
                                driver.Close();
                                driver.SwitchTo().Window(originalWindowHandle);
                                continue;
                            }
                            
                            string clickedUrl = driver.Url;
                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Đã switch sang tab mới, URL hiện tại: {clickedUrl}");
                            
                            // Kiểm tra link external
                            bool isExternal = IsExternalLink(clickedUrl);
                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Kiểm tra link external: {isExternal}");
                            
                            // Kiểm tra trang có ảnh
                            bool hasImages = CheckPageHasImages(driver);
                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Kiểm tra trang có ảnh: {hasImages}");
                            
                            if (isExternal && hasImages)
                            {
                                // Kiểm tra URL đã được xử lý chưa (tránh lặp)
                                string normalizedUrl = NormalizeUrl(clickedUrl);
                                if (processedUrls.Contains(normalizedUrl))
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚠ URL đã được xử lý trước đó, bỏ qua để tránh lặp: {normalizedUrl}");
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> URL gốc: {clickedUrl}");
                                    driver.Close();
                                    driver.SwitchTo().Window(originalWindowHandle);
                                    CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                                    continue;
                                }
                                
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✓ Trang đạt tiêu chí (external={isExternal}, hasImages={hasImages}), bắt đầu xử lý chụp screenshot...");
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> URL: {clickedUrl} (normalized: {normalizedUrl})");
                                
                                // Đánh dấu URL đã xử lý (trước khi chụp để tránh chụp lại nếu có lỗi)
                                processedUrls.Add(normalizedUrl);
                                
                                // Lấy jpegQuality từ job params, mặc định 70
                                long jpegQuality = 70;
                                try
                                {
                                    if (jobParamsSnapshot != null && jobParamsSnapshot["quanlity_image"] != null)
                                    {
                                        jpegQuality = jobParamsSnapshot["quanlity_image"].ToObject<long>();
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Lấy jpegQuality từ job params: {jpegQuality}");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Không có jpegQuality trong job params, dùng mặc định: {jpegQuality}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Lỗi khi lấy jpegQuality: {ex.Message}, dùng mặc định: {jpegQuality}");
                                }
                                
                                // Đóng popup trước khi chụp screenshot
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Bước 1: Đóng popup nếu có...");
                                CloseCommonPopups(driver);
                                System.Threading.Thread.Sleep(200);
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Đã đóng popup xong");
                                
                                // Kiểm tra xem có phải là error page không trước khi chụp
                                if (IsErrorPage(driver))
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚠ Phát hiện error page tại URL: {clickedUrl}, bỏ qua không chụp screenshot");
                                    driver.Close();
                                    driver.SwitchTo().Window(originalWindowHandle);
                                    CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                                    continue;
                                }
                                
                                // Chụp toàn bộ màn hình và chuyển sang base64 cho TẤT CẢ các link
                                string screenshotBase64 = "";
                                string contactInfoJson = "{}";
                                try
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ========== BẮT ĐẦU CHỤP FULL PAGE SCREENSHOT ==========");
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> URL: {clickedUrl}");
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> JPEG Quality: {jpegQuality}");
                                    
                                    screenshotBase64 = CaptureFullPageScreenshotAsBase64(driver, out contactInfoJson, jpegQuality);
                                    
                                    if (!string.IsNullOrWhiteSpace(screenshotBase64))
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✓✓✓ CHỤP SCREENSHOT THÀNH CÔNG ✓✓✓");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Kích thước base64: {screenshotBase64.Length} ký tự");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Contact info extracted: {contactInfoJson?.Length ?? 0} chars");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ========== HOÀN THÀNH CHỤP SCREENSHOT ==========");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✗✗✗ CHỤP SCREENSHOT THẤT BẠI - KẾT QUẢ RỖNG ✗✗✗");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ========== KẾT THÚC CHỤP SCREENSHOT (THẤT BẠI) ==========");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✗✗✗ EXCEPTION KHI CHỤP SCREENSHOT ✗✗✗");
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Lỗi: {ex.Message}");
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> StackTrace: {ex.StackTrace}");
                                    ErrorWriter.WriteLog(LogPath, "CaptureScreenshotForBanner", $"{clickedUrl} => {ex}");
                                }
                                
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Bước cuối: Push vào queue với screenshotBase64 length: {(screenshotBase64?.Length ?? 0)}");
                                if (!string.IsNullOrWhiteSpace(screenshotBase64))
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚡⚡⚡ GỌI PushIframeToQueue - clickedUrl: '{clickedUrl}', screenshotLen: {(screenshotBase64?.Length ?? 0)} ⚡⚡⚡");
                                    PushIframeToQueue(clickedUrl, screenshotBase64, jobParamsSnapshot, contactInfoJson ?? "{}");
                                processedCount++;
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✓ Đã push vào queue, processedCount: {processedCount}");
                                }
                                else
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚠ Screenshot rỗng, bỏ qua không push vào queue");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✗ Trang KHÔNG đạt tiêu chí, bỏ qua");
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} ->   - external: {isExternal}");
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} ->   - hasImages: {hasImages}");
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} ->   - URL: {clickedUrl}");
                            }
                            
                            driver.Close();
                            driver.SwitchTo().Window(originalWindowHandle);
                            CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                            System.Threading.Thread.Sleep(300);
                        }
                        else
                        {
                            if (TryWaitForUrlChange(driver, originalUrl, TimeSpan.FromSeconds(5), out var currentUrl) && currentUrl != originalUrl)
                            {
                                if (!WaitForPageLoad(driver, pageLoadTimeout))
                                {
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Trang mới load quá {pageLoadTimeout.TotalSeconds}s, quay lại trang gốc");
                                }
                                else if (IsExternalLink(currentUrl) && CheckPageHasImages(driver))
                                {
                                    // Kiểm tra URL đã được xử lý chưa (tránh lặp)
                                    string normalizedUrl = NormalizeUrl(currentUrl);
                                    if (processedUrls.Contains(normalizedUrl))
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚠ URL đã được xử lý trước đó, bỏ qua để tránh lặp: {normalizedUrl}");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> URL gốc: {currentUrl}");
                                        driver.Navigate().GoToUrl(originalUrl);
                                        System.Threading.Thread.Sleep(500);
                                        CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                                        continue;
                                    }
                                    
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✓ Trang đạt tiêu chí (navigate trong cùng tab), bắt đầu xử lý chụp screenshot...");
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> URL: {currentUrl} (normalized: {normalizedUrl})");
                                    
                                    // Đánh dấu URL đã xử lý (trước khi chụp để tránh chụp lại nếu có lỗi)
                                    processedUrls.Add(normalizedUrl);
                                    
                                    // Lấy jpegQuality từ job params, mặc định 70
                                    long jpegQuality = 70;
                                    try
                                    {
                                        if (jobParamsSnapshot != null && jobParamsSnapshot["quanlity_image"] != null)
                                        {
                                            jpegQuality = jobParamsSnapshot["quanlity_image"].ToObject<long>();
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Lấy jpegQuality từ job params: {jpegQuality}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Không có jpegQuality trong job params, dùng mặc định: {jpegQuality}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Lỗi khi lấy jpegQuality: {ex.Message}, dùng mặc định: {jpegQuality}");
                                    }
                                    
                                    // Đóng popup trước khi chụp screenshot
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Bước 1: Đóng popup nếu có...");
                                    CloseCommonPopups(driver);
                                    System.Threading.Thread.Sleep(200);
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Đã đóng popup xong");
                                    
                                    // Kiểm tra xem có phải là error page không trước khi chụp
                                    if (IsErrorPage(driver))
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚠ Phát hiện error page tại URL: {currentUrl}, bỏ qua không chụp screenshot");
                                        driver.Navigate().GoToUrl(originalUrl);
                                        System.Threading.Thread.Sleep(500);
                                        CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                                        continue;
                                    }
                                    
                                    // Chụp toàn bộ màn hình và chuyển sang base64 cho TẤT CẢ các link
                                    string screenshotBase64 = "";
                                    string contactInfoJson = "{}";
                                    try
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ========== BẮT ĐẦU CHỤP FULL PAGE SCREENSHOT (navigate) ==========");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> URL: {currentUrl}");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> JPEG Quality: {jpegQuality}");
                                        
                                        screenshotBase64 = CaptureFullPageScreenshotAsBase64(driver, out contactInfoJson, jpegQuality);
                                        
                                        if (!string.IsNullOrWhiteSpace(screenshotBase64))
                                        {
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✓✓✓ CHỤP SCREENSHOT THÀNH CÔNG ✓✓✓");
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Kích thước base64: {screenshotBase64.Length} ký tự");
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Contact info extracted: {contactInfoJson?.Length ?? 0} chars");
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ========== HOÀN THÀNH CHỤP SCREENSHOT ==========");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✗✗✗ CHỤP SCREENSHOT THẤT BẠI - KẾT QUẢ RỖNG ✗✗✗");
                                            Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ========== KẾT THÚC CHỤP SCREENSHOT (THẤT BẠI) ==========");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✗✗✗ EXCEPTION KHI CHỤP SCREENSHOT ✗✗✗");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Lỗi: {ex.Message}");
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> StackTrace: {ex.StackTrace}");
                                        ErrorWriter.WriteLog(LogPath, "CaptureScreenshotForBanner", $"{currentUrl} => {ex}");
                                    }
                                    
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Bước cuối: Push vào queue với screenshotBase64 length: {(screenshotBase64?.Length ?? 0)}");
                                    if (!string.IsNullOrWhiteSpace(screenshotBase64))
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚡⚡⚡ GỌI PushIframeToQueue - currentUrl: '{currentUrl}', screenshotLen: {(screenshotBase64?.Length ?? 0)} ⚡⚡⚡");
                                        PushIframeToQueue(currentUrl, screenshotBase64, jobParamsSnapshot, contactInfoJson ?? "{}");
                                    processedCount++;
                                    Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ✓ Đã push vào queue, processedCount: {processedCount}");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> ⚠ Screenshot rỗng, bỏ qua không push vào queue");
                                    }
                                }
                                
                                driver.Navigate().GoToUrl(originalUrl);
                                System.Threading.Thread.Sleep(500);
                                CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                            }
                            else
                            {
                                Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Click không mở tab mới hoặc navigate, bỏ qua");
                            }
                        }
                    }
                    catch (StaleElementReferenceException ex)
                    {
                        shouldMarkFingerprint = false;
                        Console.WriteLine($"[AdLink] Banner {handledBannerOrdinal} -> Element bị stale, sẽ thu thập lại: {ex.Message}");
                        System.Threading.Thread.Sleep(200);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[AdLink] Lỗi khi xử lý banner {handledBannerOrdinal}: {ex.Message}");
                        ErrorWriter.WriteLog(LogPath, "ProcessIframesAndPushToQueue", $"{hostLabel} => Banner {handledBannerOrdinal} => {ex}");
                        
                        try
                        {
                            driver.SwitchTo().Window(originalWindowHandle);
                        }
                        catch
                        {
                            if (driver.WindowHandles.Count > 0)
                            {
                                driver.SwitchTo().Window(driver.WindowHandles[0]);
                                originalWindowHandle = driver.CurrentWindowHandle;
                            }
                        }

                        CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                    }
                    finally
                    {
                        if (shouldMarkFingerprint && !string.IsNullOrWhiteSpace(bannerFingerprint))
                        {
                            processedFingerprints.Add(bannerFingerprint);
                        }
                    }
                }
                
                // Đảm bảo đang ở tab gốc và trang gốc không bị thay đổi
                try
                {
                    driver.SwitchTo().Window(originalWindowHandle);
                    // Kiểm tra xem URL có thay đổi không, nếu có thì quay lại
                    if (driver.Url != originalUrl)
                    {
                        Console.WriteLine($"[AdLink] Phát hiện URL gốc đã thay đổi, quay lại URL gốc");
                        driver.Navigate().GoToUrl(originalUrl);
                        System.Threading.Thread.Sleep(1000);
                    }
                }
                catch
                {
                    // Nếu tab gốc không còn, tìm lại tab đầu tiên
                    if (driver.WindowHandles.Count > 0)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[0]);
                        originalWindowHandle = driver.CurrentWindowHandle;
                    }
                }

                CloseAllNonOriginalTabs(driver, ref originalWindowHandle);
                
                Console.WriteLine($"[AdLink] Hoàn thành xử lý {processedCount}/{handledBannerOrdinal} banner(s) có ảnh và external cho host={hostLabel}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] Lỗi xử lý iframe: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "ProcessIframesAndPushToQueue", $"{hostLabel} => {ex}");
                TelegramService.PushLogToTelegram($"ProcessIframesAndPushToQueue Error - {hostLabel}", ex);
            }
        }
        
        /// <summary>
        /// Kiểm tra xem trang hiện tại có chứa ít nhất 1 ảnh không
        /// </summary>
        private static bool CheckPageHasImages(IWebDriver driver)
        {
            try
            {
                // Kiểm tra bằng JavaScript
                var js = (IJavaScriptExecutor)driver;
                
                // 1. Đếm số lượng thẻ <img> có src hợp lệ
                var imgCount = (long)js.ExecuteScript(@"
                    var imgs = document.querySelectorAll('img[src]');
                    var count = 0;
                    for (var i = 0; i < imgs.length; i++) {
                        var src = imgs[i].src;
                        if (src && src.trim() !== '' && !src.startsWith('data:') && src !== 'about:blank') {
                            count++;
                        }
                    }
                    return count;
                ");
                
                if (imgCount > 0)
                {
                    Console.WriteLine($"[AdLink] Tìm thấy {imgCount} ảnh từ thẻ <img>");
                    return true;
                }
                
                // 2. Kiểm tra background-image trong CSS
                var bgImageCount = (long)js.ExecuteScript(@"
                    var elements = document.querySelectorAll('*');
                    var count = 0;
                    for (var i = 0; i < elements.length; i++) {
                        var style = window.getComputedStyle(elements[i]);
                        var bgImage = style.backgroundImage;
                        if (bgImage && bgImage !== 'none' && bgImage.includes('url(') && !bgImage.includes('data:')) {
                            count++;
                        }
                    }
                    return count;
                ");
                
                if (bgImageCount > 0)
                {
                    Console.WriteLine($"[AdLink] Tìm thấy {bgImageCount} element có background-image");
                    return true;
                }
                
                // 3. Kiểm tra trong HTML source có từ khóa liên quan đến ảnh
                var htmlSource = driver.PageSource?.ToLowerInvariant() ?? "";
                if (htmlSource.Contains("<img") || 
                    htmlSource.Contains("background-image") || 
                    htmlSource.Contains(".jpg") || 
                    htmlSource.Contains(".jpeg") || 
                    htmlSource.Contains(".png") || 
                    htmlSource.Contains(".gif") || 
                    htmlSource.Contains(".webp") ||
                    htmlSource.Contains("image") ||
                    htmlSource.Contains("picture"))
                {
                    Console.WriteLine($"[AdLink] Tìm thấy dấu hiệu ảnh trong HTML source");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] Lỗi khi kiểm tra ảnh: {ex.Message}");
                // Nếu có lỗi, giả định là có ảnh để an toàn
                return true;
            }
        }

        private static bool IsFacebookLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;
            var lower = url.ToLowerInvariant();
            return lower.Contains("facebook.com") || lower.Contains("fb.com") || lower.Contains("fb.me");
        }
        
        /// <summary>
        /// Thử đóng popup bằng cách click vào nút close (icon X hoặc class = close)
        /// </summary>
        private static void TryClosePopup(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var script = @"
                    (function() {
                        // Tìm các nút close có icon X hoặc class chứa 'close'
                        var closeSelectors = [
                            'button[aria-label*=""close""]',
                            'button[aria-label*=""Close""]',
                            'button[aria-label*=""Đóng""]',
                            '.close', '.Close', '.close-button', '.popup-close', '.btn-close',
                            '[class*=""close""]', '[class*=""Close""]',
                            'button.close', 'span.close', 'div.close', 'a.close'
                        ];
                        
                        for (var i = 0; i < closeSelectors.length; i++) {
                            try {
                                var elements = document.querySelectorAll(closeSelectors[i]);
                                for (var j = 0; j < elements.length; j++) {
                                    var el = elements[j];
                                    var style = window.getComputedStyle(el);
                                    if (style && style.display !== 'none' && style.visibility !== 'hidden') {
                                        var rect = el.getBoundingClientRect();
                                        // Chỉ click nếu element visible và có kích thước hợp lý
                                        if (rect.width > 0 && rect.height > 0) {
                                            el.click();
                                            return true; // Đã click thành công
                                        }
                                    }
                                }
                            } catch(e) {}
                        }
                        return false; // Không tìm thấy nút close
                    })();
                ";
                var clicked = js.ExecuteScript(script);
                if (clicked != null && Convert.ToBoolean(clicked))
                {
                    Console.WriteLine($"[AdLink] [TryClosePopup] Đã click vào nút close để đóng popup");
                    System.Threading.Thread.Sleep(300); // Đợi popup đóng
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [TryClosePopup] Lỗi khi đóng popup: {ex.Message}");
            }
        }

        /// <summary>
        /// Tìm email hoặc số điện thoại trên trang và scroll viewport tới vị trí phần tử đó nếu tìm thấy
        /// </summary>
        private static bool TryScrollToContactInfo(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var result = js.ExecuteScript(@"
                    (function() {
                        const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
                        const phoneRegex = /(\+?\d{1,3}[\s\-\.]?)?(\(?\d{2,4}\)?[\s\-\.]?){2,4}\d{2,4}/;
                        const mstRegex = /(MST|Mã số thuế|mã số thuế)[\s:]*\d{10,13}/i;
                        const addressKeywords = ['địa chỉ', 'address', 'location', 'văn phòng', 'trụ sở', 'office'].map(k => k.toLowerCase());
                        const companyKeywords = ['công ty', 'company', 'cty', 'cty.', 'doanh nghiệp', 'enterprise'].map(k => k.toLowerCase());
                        const selectors = 'a, span, div, p, li, strong, b, h1, h2, h3, h4, h5, h6, footer, section, address';
                        const nodes = Array.from(document.querySelectorAll(selectors));
                        const scrollY = window.scrollY || window.pageYOffset || document.documentElement.scrollTop || 0;

                        for (const node of nodes) {
                            if (!node || !node.innerText) continue;
                            const text = node.innerText.trim();
                            if (!text) continue;
                            if (text.length > 200) continue;
                            
                            const textLower = text.toLowerCase();
                            let hasContactInfo = false;
                            
                            // Kiểm tra email
                            if (emailRegex.test(text)) {
                                hasContactInfo = true;
                            }
                            // Kiểm tra phone
                            else if (phoneRegex.test(text)) {
                                hasContactInfo = true;
                            }
                            // Kiểm tra MST (Mã số thuế)
                            else if (mstRegex.test(text)) {
                                hasContactInfo = true;
                            }
                            // Kiểm tra địa chỉ (address keywords)
                            else {
                                for (let i = 0; i < addressKeywords.length; i++) {
                                    if (textLower.includes(addressKeywords[i])) {
                                        hasContactInfo = true;
                                        break;
                                    }
                                }
                                // Kiểm tra công ty (company keywords)
                                if (!hasContactInfo) {
                                    for (let i = 0; i < companyKeywords.length; i++) {
                                        if (textLower.includes(companyKeywords[i])) {
                                            hasContactInfo = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            if (hasContactInfo) {
                                const rect = node.getBoundingClientRect();
                                return {
                                    found: true,
                                    top: rect.top + scrollY,
                                    height: rect.height,
                                    text: text.substring(0, 120)
                                };
                            }
                        }
                        return { found: false };
                    })();
                ") as IDictionary<string, object>;

                if (result != null && result.TryGetValue("found", out var foundObj) && Convert.ToBoolean(foundObj))
                {
                    double targetTop = 0;
                    if (result.TryGetValue("top", out var topObj))
                    {
                        double.TryParse(Convert.ToString(topObj), out targetTop);
                    }
                    double offset = 120;
                    var finalY = Math.Max(0, targetTop - offset);
                    js.ExecuteScript("window.scrollTo(0, arguments[0]);", finalY);
                    if (result.TryGetValue("text", out var textObj))
                    {
                        Console.WriteLine($"[AdLink] [TryScrollToContactInfo] Đã tìm thấy contact info (email/phone/address/company/MST): {textObj}");
                    }
                    System.Threading.Thread.Sleep(300);
                    return true;
                }

                Console.WriteLine("[AdLink] [TryScrollToContactInfo] Không tìm thấy email/số điện thoại/địa chỉ/công ty/MST, giữ nguyên logic cũ");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [TryScrollToContactInfo] Lỗi: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Scroll viewport tới sát cuối trang để đảm bảo footer hiển thị trên màn hình
        /// </summary>
        private static void ScrollViewportToBottom(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var totalHeight = Convert.ToDouble(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight, document.body.offsetHeight, document.documentElement.offsetHeight, document.body.clientHeight, document.documentElement.clientHeight) || 0;"));
                var viewportHeight = Convert.ToDouble(js.ExecuteScript("return window.innerHeight || document.documentElement.clientHeight || document.body.clientHeight || 0;"));
                var targetY = Math.Max(0, totalHeight - viewportHeight);
                js.ExecuteScript("window.scrollTo(0, arguments[0]);", targetY);
                System.Threading.Thread.Sleep(400);
                try { js.ExecuteScript("window.scrollBy(0, -20);"); } catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [ScrollViewportToBottom] Lỗi khi scroll xuống cuối trang: {ex.Message}");
            }
        }

        /// <summary>
        /// Đảm bảo viewport thực sự đang ở sát cuối trang (scrollY + viewportHeight ~ totalHeight)
        /// </summary>
        private static bool EnsureViewportAtBottom(IWebDriver driver, int retries = 3, int settleDelayMs = 300)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                for (int attempt = 1; attempt <= retries; attempt++)
                {
                    double totalHeight = Convert.ToDouble(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
                    double viewportHeight = Convert.ToDouble(js.ExecuteScript("return window.innerHeight || document.documentElement.clientHeight || 0;"));
                    double scrollY = Convert.ToDouble(js.ExecuteScript("return window.scrollY || window.pageYOffset || document.documentElement.scrollTop || 0;"));

                    if (scrollY + viewportHeight >= totalHeight - 5)
                    {
                        Console.WriteLine($"[AdLink] [EnsureViewportAtBottom] Đã ở sát cuối trang (attempt {attempt})");
                        return true;
                    }

                    var targetY = Math.Max(0, totalHeight - viewportHeight);
                    Console.WriteLine($"[AdLink] [EnsureViewportAtBottom] Chưa tới cuối trang (attempt {attempt}), scrollY={scrollY}, total={totalHeight}. Đang scroll tới {targetY}");
                    js.ExecuteScript("window.scrollTo(0, arguments[0]);", targetY);
                    System.Threading.Thread.Sleep(Math.Max(200, settleDelayMs));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [EnsureViewportAtBottom] Lỗi: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Chụp ảnh viewport hiện tại và trả về byte array
        /// </summary>
        private static byte[] CaptureCurrentViewportScreenshot(IWebDriver driver)
        {
            byte[] screenshotBytes = Array.Empty<byte>();
            var chrome = driver as ChromeDriver;
            bool success = false;

            if (chrome != null)
            {
                try
                {
                    var args = new Dictionary<string, object>
                    {
                        { "format", "jpeg" },
                        { "quality", 100 },
                        { "captureBeyondViewport", false }
                    };
                    var result = chrome.ExecuteCdpCommand("Page.captureScreenshot", args) as IDictionary<string, object>;
                    if (result != null && result.TryGetValue("data", out var dataObj) && dataObj is string base64)
                    {
                        screenshotBytes = Convert.FromBase64String(base64);
                        success = screenshotBytes != null && screenshotBytes.Length > 0;
                    }
                }
                catch { }

                if (!success)
                {
                    try
                    {
                        var shot = ((ITakesScreenshot)driver).GetScreenshot();
                        screenshotBytes = shot.AsByteArray;
                        success = screenshotBytes != null && screenshotBytes.Length > 0;
                    }
                    catch { }
                }
            }
            else
            {
                try
                {
                    var shot = ((ITakesScreenshot)driver).GetScreenshot();
                    screenshotBytes = shot.AsByteArray;
                    success = screenshotBytes != null && screenshotBytes.Length > 0;
                }
                catch { }
            }

            return success && screenshotBytes != null && screenshotBytes.Length > 0 ? screenshotBytes : Array.Empty<byte>();
        }

        /// <summary>
        /// Kiểm tra xem trang có menu/link đến /lien-he không
        /// </summary>
        private static bool HasLienHeMenu(IWebDriver driver, out string? lienHeUrl)
        {
            lienHeUrl = null;
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var currentUrl = driver.Url;
                var uri = new Uri(currentUrl);
                var baseUrl = $"{uri.Scheme}://{uri.Host}";

                var script = @"
                    (function() {
                        var links = document.querySelectorAll('a[href*=""lien-he""], a[href*=""contact""], a[href*=""lienhe""]');
                        for (var i = 0; i < links.length; i++) {
                            var href = links[i].getAttribute('href');
                            if (href && (href.includes('/lien-he') || href.includes('/contact') || href.includes('/lienhe'))) {
                                return href;
                            }
                        }
                        return null;
                    })();
                ";

                var result = js.ExecuteScript(script);
                if (result != null && !string.IsNullOrWhiteSpace(result.ToString()))
                {
                    var href = result.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(href) && (href.StartsWith("http://") || href.StartsWith("https://")))
                    {
                        lienHeUrl = href;
                    }
                    else if (href.StartsWith("/"))
                    {
                        lienHeUrl = baseUrl + href;
                    }
                    else
                    {
                        lienHeUrl = baseUrl + "/" + href;
                    }
                    Console.WriteLine($"[AdLink] [HasLienHeMenu] Tìm thấy menu /lien-he: {lienHeUrl}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [HasLienHeMenu] Lỗi khi kiểm tra menu /lien-he: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Kiểm tra xem trang hiện tại có phải là trang lỗi "This site can't be reached" hoặc các lỗi kết nối khác không
        /// </summary>
        private static bool IsErrorPage(IWebDriver driver)
        {
            try
            {
                var currentUrl = driver.Url;
                
                // Kiểm tra URL có phải là error page của browser không
                if (currentUrl.StartsWith("chrome-error://", StringComparison.OrdinalIgnoreCase) ||
                    currentUrl.StartsWith("edge-error://", StringComparison.OrdinalIgnoreCase) ||
                    currentUrl.StartsWith("about:error", StringComparison.OrdinalIgnoreCase))
            {
                    Console.WriteLine($"[AdLink] [IsErrorPage] Phát hiện error page từ URL: {currentUrl}");
                    return true;
                }
                
                // Kiểm tra title và body text
                var js = (IJavaScriptExecutor)driver;
                var result = js.ExecuteScript(@"
                    (function() {
                        var title = document.title || '';
                        var bodyText = document.body ? (document.body.innerText || document.body.textContent || '') : '';
                        var combined = (title + ' ' + bodyText).toLowerCase();
                
                        // Các từ khóa lỗi phổ biến
                        var errorKeywords = [
                            'this site can\'t be reached',
                            'this site cannot be reached',
                            'err_name_not_resolved',
                            'err_connection_refused',
                            'err_internet_disconnected',
                            'err_connection_timed_out',
                            'err_connection_reset',
                            'err_network_changed',
                            'err_timed_out',
                            'unable to connect',
                            'can\'t reach this page',
                            'cannot reach this page',
                            'site unreachable',
                            'dns_probe_finished_nxdomain',
                            'net::err_name_not_resolved',
                            'net::err_connection_refused'
                        ];
                        
                        for (var i = 0; i < errorKeywords.length; i++) {
                            if (combined.includes(errorKeywords[i])) {
                                return true;
                            }
                        }
                        
                        return false;
                    })();
                ") as bool?;
                
                if (result == true)
                {
                    Console.WriteLine($"[AdLink] [IsErrorPage] Phát hiện error page từ nội dung trang");
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [IsErrorPage] Lỗi khi kiểm tra error page: {ex.Message}");
                // Nếu có lỗi khi kiểm tra, giả định không phải error page để tiếp tục xử lý
                return false;
            }
        }

        /// <summary>
        /// Loại bỏ các overlay, mask, hoặc elements có z-index cao che mất thông tin contact (TỐI ƯU)
        /// </summary>
        private static void RemoveOverlaysAndMasks(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var script = @"
                    (function() {
                        let removed = 0;
                        const winWidth = window.innerWidth;
                        const winHeight = window.innerHeight;
                        const minOverlayWidth = winWidth * 0.5;
                        const minOverlayHeight = winHeight * 0.5;
                        
                        // 1. Tìm và ẩn/xóa các overlay phổ biến (tối ưu: chỉ query các selector cụ thể)
                        const overlaySelectors = [
                            '.overlay', '.backdrop', '.modal-backdrop', '.modal-mask',
                            '.ant-modal-mask', '.ReactModal__Overlay', '.popup-overlay'
                        ];
                        
                        for (let i = 0; i < overlaySelectors.length; i++) {
                            try {
                                const elements = document.querySelectorAll(overlaySelectors[i]);
                                for (let j = 0; j < elements.length && j < 10; j++) { // Giới hạn tối đa 10 elements mỗi selector
                                    const el = elements[j];
                                    const rect = el.getBoundingClientRect();
                                    // Early return: bỏ qua nếu kích thước quá nhỏ
                                    if (rect.width < minOverlayWidth && rect.height < minOverlayHeight) continue;
                                    
                                    const style = window.getComputedStyle(el);
                                    const zIndex = parseInt(style.zIndex) || 0;
                                    // Chỉ xóa nếu là overlay thực sự (có z-index cao và che màn hình)
                                    if (zIndex > 1000) {
                                        el.style.display = 'none';
                                        removed++;
                                    }
                                }
                            } catch(e) {}
                        }
                        
                        // 2. Tìm elements có z-index cao (tối ưu: chỉ query elements có position fixed/absolute)
                        // Chỉ xử lý elements có position fixed hoặc absolute để giảm số lượng
                        const positionedElements = document.querySelectorAll('[style*=""position:fixed""], [style*=""position:absolute""]');
                        const maxElements = Math.min(positionedElements.length, 50); // Giới hạn tối đa 50 elements
                        
                        for (let i = 0; i < maxElements; i++) {
                            try {
                                const el = positionedElements[i];
                                const rect = el.getBoundingClientRect();
                                // Early return: bỏ qua nếu không ở phần dưới màn hình
                                if (rect.top >= winHeight * 0.7) continue;
                                if (rect.width < winWidth * 0.3 && rect.height < winHeight * 0.3) continue;
                                
                                const style = window.getComputedStyle(el);
                                const zIndex = parseInt(style.zIndex) || 0;
                
                                // Nếu element có z-index > 1000 và che mất phần dưới trang
                                if (zIndex > 1000 && style.position !== 'static') {
                                    // Kiểm tra xem có phải là popup/modal không (tối ưu: kiểm tra class trước)
                                    const className = el.className || '';
                                    const role = el.getAttribute('role') || '';
                                    const isPopup = className.includes('modal') || 
                                                  className.includes('popup') ||
                                                  role === 'dialog' ||
                                                  role === 'alertdialog';
                                    
                                    if (!isPopup) {
                                        el.style.display = 'none';
                                        removed++;
                                    }
                                }
                            } catch(e) {}
                        }
                        
                        // 3. Loại bỏ fixed/sticky elements che mất footer (tối ưu: chỉ query fixed elements)
                        const fixedElements = document.querySelectorAll('[style*=""position:fixed""]');
                        const maxFixed = Math.min(fixedElements.length, 20); // Giới hạn tối đa 20 elements
                        
                        for (let i = 0; i < maxFixed; i++) {
                            try {
                                const el = fixedElements[i];
                                const rect = el.getBoundingClientRect();
                                // Early return: bỏ qua nếu không ở cuối trang
                                if (rect.top + rect.height <= winHeight * 0.8) continue;
                                if (rect.height <= 50) continue;
                                
                                el.style.opacity = '0.3'; // Làm mờ thay vì ẩn hoàn toàn
                                removed++;
                            } catch(e) {}
                        }
                        
                        return removed;
                    })();
                ";
                
                var removed = js.ExecuteScript(script);
                if (removed != null && Convert.ToInt32(removed) > 0)
                {
                    Console.WriteLine($"[AdLink] [RemoveOverlays] Đã loại bỏ {removed} overlay/mask che mất thông tin");
                    System.Threading.Thread.Sleep(100); // Giảm từ 500ms xuống 100ms
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [RemoveOverlays] Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Tìm và hiển thị các elements bị ẩn nhưng chứa thông tin contact (TỐI ƯU)
        /// </summary>
        private static void RevealHiddenContactInfo(IWebDriver driver)
        {
                try
                {
                var js = (IJavaScriptExecutor)driver;
                var script = @"
                    (function() {
                        const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
                        const phoneRegex = /(\+?\d{1,3}[\s\-\.]?)?(\(?\d{2,4}\)?[\s\-\.]?){2,4}\d{2,4}/;
                        const mstRegex = /(MST|Mã số thuế|mã số thuế)[\s:]*\d{10,13}/i;
                        const addressKeywords = ['địa chỉ', 'address', 'location', 'văn phòng', 'trụ sở'].map(k => k.toLowerCase());
                        const companyKeywords = ['công ty', 'company', 'cty', 'cty.', 'doanh nghiệp', 'enterprise'].map(k => k.toLowerCase());
                        
                        let revealed = 0;
                        // Tối ưu: Chỉ query các elements có khả năng chứa contact info (div, span, p, a, footer, section)
                        const candidateSelectors = 'div, span, p, a, footer, section, address, li';
                        const allElements = document.querySelectorAll(candidateSelectors);
                        const maxElements = Math.min(allElements.length, 200); // Giới hạn tối đa 200 elements
                        
                        for (let i = 0; i < maxElements; i++) {
                            try {
                                const el = allElements[i];
                                // Early return: bỏ qua nếu element quá lớn (không phải contact info)
                                const text = el.innerText || el.textContent || '';
                                if (!text || text.length > 500 || text.length < 5) continue;
                                
                                // Tối ưu: kiểm tra text trước khi gọi getComputedStyle (tốn kém hơn)
                                const textLower = text.toLowerCase();
                                let hasContactInfo = false;
                                
                                // Kiểm tra email
                                if (emailRegex.test(text)) {
                                    hasContactInfo = true;
                                }
                                // Kiểm tra phone (chỉ nếu chưa tìm thấy email)
                                else if (phoneRegex.test(text)) {
                                    hasContactInfo = true;
                                }
                                // Kiểm tra MST (Mã số thuế)
                                else if (mstRegex.test(text)) {
                                    hasContactInfo = true;
                                }
                                // Kiểm tra keywords (chỉ nếu chưa tìm thấy)
                                else {
                                    for (let k = 0; k < addressKeywords.length; k++) {
                                        if (textLower.includes(addressKeywords[k])) {
                                            hasContactInfo = true;
                                            break;
                                        }
                                    }
                                    // Kiểm tra công ty (chỉ nếu chưa tìm thấy)
                                    if (!hasContactInfo) {
                                        for (let k = 0; k < companyKeywords.length; k++) {
                                            if (textLower.includes(companyKeywords[k])) {
                                                hasContactInfo = true;
                                                break;
                                            }
                                        }
                                    }
                                }
                                
                                if (!hasContactInfo) continue;
                                
                                // Chỉ gọi getComputedStyle khi đã xác nhận có contact info
                                const style = window.getComputedStyle(el);
                                const rect = el.getBoundingClientRect();
                                
                                // Kiểm tra nếu element bị ẩn
                                const isHidden = style.display === 'none' || 
                                               style.visibility === 'hidden' ||
                                               parseFloat(style.opacity) === 0 ||
                                               rect.width === 0 || 
                                               rect.height === 0;
                                
                                if (isHidden) {
                                    // Hiển thị element
                                    el.style.display = 'block';
                                    el.style.visibility = 'visible';
                                    el.style.opacity = '1';
                                    el.style.height = 'auto';
                                    el.style.width = 'auto';
                                    
                                    // Đảm bảo parent elements cũng hiển thị (giới hạn depth)
                                    let parent = el.parentElement;
                                    let depth = 0;
                                    while (parent && depth < 2) { // Giảm từ 3 xuống 2
                                        const parentStyle = window.getComputedStyle(parent);
                                        if (parentStyle.display === 'none') {
                                            parent.style.display = 'block';
                                        }
                                        if (parentStyle.visibility === 'hidden') {
                                            parent.style.visibility = 'visible';
                                        }
                                        parent = parent.parentElement;
                                        depth++;
                                    }
                                    
                                    revealed++;
                                    // Giới hạn số lượng elements được reveal để tránh quá chậm
                                    if (revealed >= 10) break;
                    }
                            } catch(e) {}
                        }
                        
                        return revealed;
                    })();
                ";
                
                var revealed = js.ExecuteScript(script);
                if (revealed != null && Convert.ToInt32(revealed) > 0)
                    {
                    Console.WriteLine($"[AdLink] [RevealHiddenContact] Đã hiển thị {revealed} elements chứa contact info");
                    System.Threading.Thread.Sleep(100); // Giảm từ 500ms xuống 100ms
                    }
                }
                catch (Exception ex)
                {
                Console.WriteLine($"[AdLink] [RevealHiddenContact] Lỗi: {ex.Message}");
            }
        }

        /// <summary>
        /// Trích xuất text contact info từ DOM (kể cả khi bị ẩn) và thêm vào payload
        /// </summary>
        private static string ExtractContactInfoFromDOM(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var script = @"
                    (function() {
                        const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
                        const phoneRegex = /(\+?\d{1,3}[\s\-\.]?)?(\(?\d{2,4}\)?[\s\-\.]?){2,4}\d{2,4}/;
                        const mstRegex = /(MST|Mã số thuế|mã số thuế)[\s:]*(\d{10,13})/i;
                        const addressKeywords = ['địa chỉ', 'address', 'location', 'văn phòng', 'trụ sở', 'office'].map(k => k.toLowerCase());
                        const companyKeywords = ['công ty', 'company', 'cty', 'cty.', 'doanh nghiệp', 'enterprise'].map(k => k.toLowerCase());
                        
                        const results = {
                            emails: [],
                            phones: [],
                            addresses: [],
                            companies: [],
                            msts: []
                        };
                        
                        const allElements = document.querySelectorAll('*');
                        allElements.forEach(el => {
                            try {
                                // Lấy text từ innerText, textContent, và cả value attribute
                                const texts = [
                                    el.innerText || '',
                                    el.textContent || '',
                                    el.getAttribute('value') || '',
                                    el.getAttribute('data-value') || '',
                                    el.getAttribute('placeholder') || ''
                                ].filter(t => t && t.length > 0 && t.length < 500);
                                
                                texts.forEach(text => {
                                    // Tìm email
                                    const emailMatches = text.match(emailRegex);
                                    if (emailMatches) {
                                        emailMatches.forEach(email => {
                                            if (!results.emails.includes(email)) {
                                                results.emails.push(email);
                                            }
                                        });
                                    }
                                    
                                    // Tìm số điện thoại
                                    const phoneMatches = text.match(phoneRegex);
                                    if (phoneMatches) {
                                        phoneMatches.forEach(phone => {
                                            const cleanPhone = phone.trim();
                                            if (cleanPhone.length >= 8 && !results.phones.includes(cleanPhone)) {
                                                results.phones.push(cleanPhone);
                                            }
                                        });
                                    }
                                    
                                    // Tìm địa chỉ (dòng chứa keyword + context)
                                    addressKeywords.forEach(keyword => {
                                        if (text.toLowerCase().includes(keyword)) {
                                            // Lấy 100 ký tự xung quanh keyword
                                            const index = text.toLowerCase().indexOf(keyword);
                                            const start = Math.max(0, index - 50);
                                            const end = Math.min(text.length, index + keyword.length + 50);
                                            const addressSnippet = text.substring(start, end).trim();
                                            
                                            if (addressSnippet.length > 10 && !results.addresses.includes(addressSnippet)) {
                                                results.addresses.push(addressSnippet);
                                            }
                                        }
                                    });
                                    
                                    // Tìm MST (Mã số thuế)
                                    const mstMatches = text.match(mstRegex);
                                    if (mstMatches && mstMatches.length >= 3) {
                                        const mstValue = mstMatches[2];
                                        if (mstValue && !results.msts.includes(mstValue)) {
                                            results.msts.push(mstValue);
                                        }
                                    }
                                    
                                    // Tìm thông tin công ty (dòng chứa keyword + context)
                                    const textLower = text.toLowerCase();
                                    companyKeywords.forEach(keyword => {
                                        if (textLower.includes(keyword)) {
                                            // Lấy 100 ký tự xung quanh keyword
                                            const index = textLower.indexOf(keyword);
                                            const start = Math.max(0, index - 50);
                                            const end = Math.min(text.length, index + keyword.length + 50);
                                            const companySnippet = text.substring(start, end).trim();
                                            
                                            if (companySnippet.length > 10 && !results.companies.includes(companySnippet)) {
                                                results.companies.push(companySnippet);
                                            }
                                        }
                                    });
                                });
                            } catch(e) {}
                        });
                        
                        return JSON.stringify(results);
                    })();
                ";
                
                var resultJson = js.ExecuteScript(script) as string;
                if (!string.IsNullOrWhiteSpace(resultJson))
                {
                    Console.WriteLine($"[AdLink] [ExtractContactInfo] Đã trích xuất contact info từ DOM");
                    return resultJson;
                }
            }
            catch (Exception ex)
                    {
                Console.WriteLine($"[AdLink] [ExtractContactInfo] Lỗi: {ex.Message}");
            }
            return "{}";
        }

        /// <summary>
        /// Đợi tất cả animations và transitions hoàn tất
        /// </summary>
        private static void WaitForAnimations(IWebDriver driver)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                var script = @"
                    (function() {
                        return new Promise((resolve) => {
                            // Đợi tất cả animations
                            const animations = document.getAnimations();
                            if (animations.length === 0) {
                                resolve();
                                return;
                            }
                            
                            Promise.all(animations.map(anim => anim.finished)).then(() => {
                                // Đợi thêm một chút để đảm bảo
                                setTimeout(resolve, 300);
                            });
                        });
                    })();
                ";
                
                try
                {
                    js.ExecuteAsyncScript(script);
                }
                catch
                {
                    // Fallback nếu async script không hoạt động
                }
                System.Threading.Thread.Sleep(500); // Fallback
            }
            catch
            {
                System.Threading.Thread.Sleep(1000); // Fallback nếu async script không hoạt động
            }
        }

        /// <summary>
        /// Chụp toàn bộ màn hình trang web và chuyển sang base64 string
        /// Quy trình mới: Chụp 3 ảnh và ghép lại
        /// 1. Chụp ảnh đầu tiên khi vào trang (đầu trang)
        /// 2. Scroll tới cuối trang và chụp ảnh thứ 2
        /// 3. Kiểm tra có menu /lien-he không, nếu có thì chuyển sang trang đó và chụp ảnh thứ 3
        /// 4. Ghép 3 ảnh lại theo chiều dọc và convert sang base64
        /// 
        /// Tham số:
        /// - driver: WebDriver
        /// - contactInfoJson: Output parameter chứa contact info trích xuất từ DOM
        /// - jpegQuality: Chất lượng JPEG (1-100, mặc định 80)
        /// 
        /// Trả về: Base64 string của ảnh đã ghép, hoặc chuỗi rỗng nếu lỗi
        /// </summary>
        private static string CaptureFullPageScreenshotAsBase64(IWebDriver driver, out string contactInfoJson, long jpegQuality = 80)
        {
            contactInfoJson = "{}";
            string? originalUrl = null;
            try
                    {
                Console.WriteLine($"[AdLink] [CaptureFullPage] Bắt đầu quy trình chụp ảnh và ghép lại...");
                
                // Kiểm tra xem có phải là error page không
                if (IsErrorPage(driver))
                {
                    Console.WriteLine($"[AdLink] [CaptureFullPage] ⚠ Phát hiện error page, bỏ qua không chụp screenshot");
                    return "";
                }
                
                var js = (IJavaScriptExecutor)driver;
                originalUrl = driver.Url;
                
                var screenshots = new List<byte[]>();
                
                // ========== BƯỚC 0: KIỂM TRA VÀ ƯU TIÊN CHỤP TRANG /lien-he TRƯỚC ==========
                Console.WriteLine($"[AdLink] [CaptureFullPage] Bước 0: Kiểm tra và ưu tiên chụp trang /lien-he...");
                
                if (HasLienHeMenu(driver, out var lienHeUrl) && !string.IsNullOrWhiteSpace(lienHeUrl))
                {
                    try
                    {
                        Console.WriteLine($"[AdLink] [CaptureFullPage] ✓ Tìm thấy menu /lien-he, ƯU TIÊN chụp trang này trước: {lienHeUrl}");
                        
                        // Navigate tới trang /lien-he
                        if (TryNavigate(driver, lienHeUrl, TimeSpan.FromSeconds(30), out var navFailReason))
                        {
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Đã navigate tới trang /lien-he thành công");
                            System.Threading.Thread.Sleep(2000);
                            
                            // Kiểm tra xem có phải là error page không
                            if (IsErrorPage(driver))
                            {
                                Console.WriteLine($"[AdLink] [CaptureFullPage] ⚠ Phát hiện error page tại /lien-he, bỏ qua không chụp");
                                // Quay lại trang gốc
                                if (!string.IsNullOrWhiteSpace(originalUrl))
                                {
                                    TryNavigate(driver, originalUrl, TimeSpan.FromSeconds(30), out _);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                // Bỏ qua phần chụp /lien-he, tiếp tục chụp trang chính
                            }
                            else
                            {
                                // Scroll lên đầu trang /lien-he
                            try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                            System.Threading.Thread.Sleep(300);
                            
                            // Loại bỏ overlays và hiển thị hidden elements ở trang /lien-he
                            RemoveOverlaysAndMasks(driver);
                            RevealHiddenContactInfo(driver);
                            WaitForAnimations(driver);
                            
                            // Thử đóng popup nếu có
                            TryClosePopup(driver);
                            System.Threading.Thread.Sleep(300);
                            
                            // Trích xuất contact info từ DOM (từ trang /lien-he - quan trọng nhất)
                            var contactInfoLienHe = ExtractContactInfoFromDOM(driver);
                            if (!string.IsNullOrWhiteSpace(contactInfoLienHe) && contactInfoLienHe != "{}")
                            {
                                contactInfoJson = contactInfoLienHe; // Gán trực tiếp vì đây là nguồn ưu tiên
                                Console.WriteLine($"[AdLink] [CaptureFullPage] Đã trích xuất contact info từ trang /lien-he");
                            }
                            
                            TryScrollToContactInfo(driver);
                            
                            // Chụp ảnh trang /lien-he (ưu tiên đầu tiên)
                            try
                            {
                                var lienHeBytes = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
                                if (lienHeBytes != null && lienHeBytes.Length > 0)
                                {
                                    screenshots.Add(lienHeBytes);
                                    Console.WriteLine($"[AdLink] [CaptureFullPage] ✓ Đã chụp ảnh trang /lien-he (ưu tiên), kích thước: {lienHeBytes.Length} bytes");
                                }
                                else
                                {
                                    Console.WriteLine($"[AdLink] [CaptureFullPage] WARNING: Không thể chụp ảnh trang /lien-he");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[AdLink] [CaptureFullPage] Lỗi khi chụp ảnh trang /lien-he: {ex.Message}");
                            }
                            
                            // Scroll xuống cuối trang /lien-he để chụp thêm nếu có thông tin ở cuối
                            ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(6));
                            ScrollViewportToBottom(driver);
                            System.Threading.Thread.Sleep(300);
                            
                            RemoveOverlaysAndMasks(driver);
                            RevealHiddenContactInfo(driver);
                            WaitForAnimations(driver);
                            TryClosePopup(driver);
                            System.Threading.Thread.Sleep(300);
                            
                            // Trích xuất thêm contact info từ cuối trang /lien-he
                            var contactInfoLienHeBottom = ExtractContactInfoFromDOM(driver);
                            if (!string.IsNullOrWhiteSpace(contactInfoLienHeBottom) && contactInfoLienHeBottom != "{}")
                            {
                                try
                                {
                                    var existingInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(contactInfoJson ?? "{}");
                                    var bottomInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(contactInfoLienHeBottom);
                                    if (existingInfo != null && bottomInfo != null)
                                    {
                                        foreach (var prop in bottomInfo.Properties())
                                        {
                                            if (prop.Value is Newtonsoft.Json.Linq.JArray bottomArray)
                                            {
                                                var existingArray = existingInfo[prop.Name] as Newtonsoft.Json.Linq.JArray;
                                                if (existingArray == null)
                                                {
                                                    existingInfo[prop.Name] = new Newtonsoft.Json.Linq.JArray();
                                                    existingArray = existingInfo[prop.Name] as Newtonsoft.Json.Linq.JArray;
                                                }
                                                if (existingArray != null)
                                                {
                                                    foreach (var item in bottomArray)
                                                    {
                                                        if (!existingArray.Any(x => x.ToString() == item.ToString()))
                                                        {
                                                            existingArray.Add(item);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        contactInfoJson = existingInfo.ToString(Newtonsoft.Json.Formatting.None);
                                    }
                                }
                                catch { }
                            }
                            
                            // Chụp ảnh cuối trang /lien-he
                            try
                            {
                                var lienHeBottomBytes = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
                                if (lienHeBottomBytes != null && lienHeBottomBytes.Length > 0)
                                {
                                    screenshots.Add(lienHeBottomBytes);
                                    Console.WriteLine($"[AdLink] [CaptureFullPage] ✓ Đã chụp ảnh cuối trang /lien-he, kích thước: {lienHeBottomBytes.Length} bytes");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[AdLink] [CaptureFullPage] Lỗi khi chụp ảnh cuối trang /lien-he: {ex.Message}");
                            }
                            
                                // Quay lại trang gốc để chụp tiếp
                                Console.WriteLine($"[AdLink] [CaptureFullPage] Quay lại trang gốc để chụp tiếp: {originalUrl}");
                            if (!string.IsNullOrWhiteSpace(originalUrl))
                            {
                                TryNavigate(driver, originalUrl, TimeSpan.FromSeconds(30), out _);
                                    System.Threading.Thread.Sleep(1000);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Không thể navigate tới trang /lien-he: {navFailReason}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[AdLink] [CaptureFullPage] Lỗi khi xử lý trang /lien-he: {ex.Message}");
                        // Quay lại trang gốc nếu có lỗi
                        if (!string.IsNullOrWhiteSpace(originalUrl))
                        {
                            try { TryNavigate(driver, originalUrl, TimeSpan.FromSeconds(30), out _); } catch { }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"[AdLink] [CaptureFullPage] Không tìm thấy menu /lien-he, tiếp tục chụp trang chính");
                }
                
                // ========== BƯỚC 1: Chụp ảnh đầu tiên - màn hình cơ sở 1 (đầu trang chính) ==========
                Console.WriteLine($"[AdLink] [CaptureFullPage] Bước 1: Chụp ảnh đầu trang chính...");
                
                // Scroll lên đầu trang
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                System.Threading.Thread.Sleep(300);
                
                // THÊM MỚI: Loại bỏ overlays và hiển thị hidden elements
                Console.WriteLine($"[AdLink] [CaptureFullPage] Loại bỏ overlays che mất thông tin...");
                RemoveOverlaysAndMasks(driver);
                RevealHiddenContactInfo(driver);
                
                // Đợi animations hoàn tất
                WaitForAnimations(driver);
                
                // Thử đóng popup nếu có
                TryClosePopup(driver);
                System.Threading.Thread.Sleep(300); // Đợi thêm để đảm bảo DOM đã cập nhật
                
                // Trích xuất contact info từ DOM (từ trang chính - merge vào nếu chưa có)
                var contactInfoMain = ExtractContactInfoFromDOM(driver);
                if (!string.IsNullOrWhiteSpace(contactInfoMain) && contactInfoMain != "{}")
                {
                    try
                    {
                        var existingInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(contactInfoJson ?? "{}");
                        var mainInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(contactInfoMain);
                        if (existingInfo != null && mainInfo != null)
                        {
                            foreach (var prop in mainInfo.Properties())
                            {
                                if (prop.Value is Newtonsoft.Json.Linq.JArray mainArray)
                                {
                                    var existingArray = existingInfo[prop.Name] as Newtonsoft.Json.Linq.JArray;
                                    if (existingArray == null)
                                    {
                                        existingInfo[prop.Name] = new Newtonsoft.Json.Linq.JArray();
                                        existingArray = existingInfo[prop.Name] as Newtonsoft.Json.Linq.JArray;
                                    }
                                    if (existingArray != null)
                                    {
                                        foreach (var item in mainArray)
                                        {
                                            if (!existingArray.Any(x => x.ToString() == item.ToString()))
                                            {
                                                existingArray.Add(item);
                                            }
                                        }
                                    }
                                }
                            }
                            contactInfoJson = existingInfo.ToString(Newtonsoft.Json.Formatting.None);
                        }
                    }
                    catch { }
                }
                
                // Nếu tìm thấy email/sđt thì scroll tới đó trước khi chụp
                var contactScrolled = TryScrollToContactInfo(driver);
                if (!contactScrolled)
                {
                    try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                }
                
                // Chụp ảnh đầu tiên
                try
                {
                    var topBytes = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
                    if (topBytes != null && topBytes.Length > 0)
                    {
                        screenshots.Add(topBytes);
                        Console.WriteLine($"[AdLink] [CaptureFullPage] Đã chụp ảnh đầu trang chính, kích thước: {topBytes.Length} bytes");
                    }
                    else
                    {
                        Console.WriteLine($"[AdLink] [CaptureFullPage] WARNING: Không thể chụp ảnh đầu trang chính");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[AdLink] [CaptureFullPage] Lỗi khi chụp ảnh 1: {ex.Message}");
                }
                
                // ========== BƯỚC 2: Scroll tới cuối trang chính và chụp ảnh ==========
                Console.WriteLine($"[AdLink] [CaptureFullPage] Bước 2: Scroll tới cuối trang chính và chụp ảnh...");
                
                // Scroll xuống cuối trang và đảm bảo đã load nội dung lazy
                ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(6));
                ScrollViewportToBottom(driver);
                var contactBottom = TryScrollToContactInfo(driver);
                if (!contactBottom)
                {
                    var bottomReached = EnsureViewportAtBottom(driver);
                    if (!bottomReached)
                    {
                        Console.WriteLine("[AdLink] [CaptureFullPage] Cảnh báo: chưa chắc chắn đang ở cuối trang, thử scroll thêm một lần nữa");
                        ScrollViewportToBottom(driver);
                        EnsureViewportAtBottom(driver);
                    }
                }
                System.Threading.Thread.Sleep(300);
                
                // THÊM MỚI: Loại bỏ overlays và hiển thị hidden elements ở cuối trang
                RemoveOverlaysAndMasks(driver);
                RevealHiddenContactInfo(driver);
                WaitForAnimations(driver);
                
                // Thử đóng popup nếu có
                TryClosePopup(driver);
                System.Threading.Thread.Sleep(300);
                
                // THÊM MỚI: Trích xuất contact info từ DOM (lần 2 - ở cuối trang)
                var contactInfoBottom = ExtractContactInfoFromDOM(driver);
                bool shouldRecapture = false;
                if (!string.IsNullOrWhiteSpace(contactInfoBottom) && contactInfoBottom != "{}")
                {
                    // Merge contact info từ cuối trang vào kết quả
                    try
                    {
                        var existingInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(contactInfoJson ?? "{}");
                        var bottomInfo = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(contactInfoBottom);
                        if (existingInfo != null && bottomInfo != null)
                        {
                            // Merge arrays (emails, phones, addresses)
                            foreach (var prop in bottomInfo.Properties())
                            {
                                if (prop.Value is Newtonsoft.Json.Linq.JArray bottomArray)
                                {
                                    var existingArray = existingInfo[prop.Name] as Newtonsoft.Json.Linq.JArray;
                                    if (existingArray == null)
                                    {
                                        existingInfo[prop.Name] = new Newtonsoft.Json.Linq.JArray();
                                        existingArray = existingInfo[prop.Name] as Newtonsoft.Json.Linq.JArray;
                                    }
                                    foreach (var item in bottomArray)
                                    {
                                        if (!existingArray.Any(x => x.ToString() == item.ToString()))
                                        {
                                            existingArray.Add(item);
                                        }
                                    }
                                }
                            }
                            contactInfoJson = existingInfo.ToString(Newtonsoft.Json.Formatting.None);
                            
                            // THÊM MỚI: Kiểm tra nếu có contact info (email/phone/address/company/MST) thì cần chụp lại
                            var hasEmails = (bottomInfo["emails"] as Newtonsoft.Json.Linq.JArray)?.Count > 0;
                            var hasPhones = (bottomInfo["phones"] as Newtonsoft.Json.Linq.JArray)?.Count > 0;
                            var hasAddresses = (bottomInfo["addresses"] as Newtonsoft.Json.Linq.JArray)?.Count > 0;
                            var hasCompanies = (bottomInfo["companies"] as Newtonsoft.Json.Linq.JArray)?.Count > 0;
                            var hasMsts = (bottomInfo["msts"] as Newtonsoft.Json.Linq.JArray)?.Count > 0;
                            
                            if (hasEmails || hasPhones || hasAddresses || hasCompanies || hasMsts)
                            {
                                Console.WriteLine($"[AdLink] [CaptureFullPage] Phát hiện contact info ở cuối trang - Emails: {hasEmails}, Phones: {hasPhones}, Addresses: {hasAddresses}, Companies: {hasCompanies}, MSTs: {hasMsts}");
                                Console.WriteLine($"[AdLink] [CaptureFullPage] Scroll lại tới contact info và chụp lại...");
                                shouldRecapture = true;
                            }
                        }
                    }
                    catch { }
                }
                
                // Nếu có contact info, scroll lại và chụp lại
                if (shouldRecapture)
                {
                    var contactScrolledAgain = TryScrollToContactInfo(driver);
                    if (contactScrolledAgain)
                    {
                        System.Threading.Thread.Sleep(300);
                        RemoveOverlaysAndMasks(driver);
                        RevealHiddenContactInfo(driver);
                        WaitForAnimations(driver);
                        TryClosePopup(driver);
                        System.Threading.Thread.Sleep(300);
                    }
                }
                
                // Chụp ảnh thứ 2
                try
                {
                    var bottomBytes = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
                    if (bottomBytes != null && bottomBytes.Length > 0)
                    {
                        screenshots.Add(bottomBytes);
                        Console.WriteLine($"[AdLink] [CaptureFullPage] Đã chụp ảnh cuối trang chính, kích thước: {bottomBytes.Length} bytes");
                    }
                    else
                    {
                        Console.WriteLine($"[AdLink] [CaptureFullPage] WARNING: Không thể chụp ảnh cuối trang chính");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[AdLink] [CaptureFullPage] Lỗi khi chụp ảnh 2: {ex.Message}");
                }
                
                // ========== BƯỚC 3: Ghép các ảnh lại và convert sang base64 ==========
                if (screenshots.Count == 0)
                {
                    Console.WriteLine($"[AdLink] [CaptureFullPage] ERROR: Không có ảnh nào để ghép");
                    return "";
                }
                
                Console.WriteLine($"[AdLink] [CaptureFullPage] Bước 4: Ghép {screenshots.Count} ảnh lại...");
                
                var images = new List<Image<Rgba32>>();
                try
                {
                    // Load tất cả ảnh vào memory
                    int maxWidth = 0;
                    int totalHeight = 0;
                    
                    foreach (var screenshotBytes in screenshots)
                    {
                        if (screenshotBytes == null || screenshotBytes.Length == 0) continue;
                        
                        using (var ms = new MemoryStream(screenshotBytes))
                        {
                            var img = Image.Load<Rgba32>(ms);
                            images.Add(img);
                            maxWidth = Math.Max(maxWidth, img.Width);
                            totalHeight += img.Height;
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Load ảnh: {img.Width}x{img.Height}");
                        }
                    }
                    
                    if (images.Count == 0)
                    {
                        Console.WriteLine($"[AdLink] [CaptureFullPage] ERROR: Không thể load ảnh nào");
                        return "";
                    }
                    
                    // Tạo ảnh mới với kích thước tổng hợp
                    Console.WriteLine($"[AdLink] [CaptureFullPage] Tạo ảnh ghép: {maxWidth}x{totalHeight}");
                    using (var mergedImage = new Image<Rgba32>(maxWidth, totalHeight))
                    {
                        int currentY = 0;
                        
                        // Vẽ từng ảnh lên ảnh ghép
                        foreach (var img in images)
                        {
                            mergedImage.Mutate(ctx => ctx.DrawImage(img, new Point(0, currentY), 1f));
                            currentY += img.Height;
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Đã vẽ ảnh tại Y={currentY - img.Height}");
                        }
                        
                        // Resize và nén ảnh để giảm dung lượng xuống dưới 800KB
                        // Đồng thời đảm bảo ảnh trong khoảng 384x384 đến 768x768 (tối ưu token Gemini)
                        Console.WriteLine($"[AdLink] [CaptureFullPage] Resize và nén ảnh (target: <800KB, optimize for Gemini: 384-768px)...");
                        
                        // Gemini API token optimization:
                        // - Ảnh ≤ 384x384: 258 token (1 ảnh)
                        // - Ảnh > 768x768: Gemini chia thành các ô 768x768, mỗi ô = 258 token
                        // - Tối ưu: Giữ ảnh trong khoảng 384x384 đến 768x768 để tối đa 258 token/ảnh
                        const int minDimension = 384;  // Giữ chất lượng tối thiểu
                        const int maxDimension = 768;  // Tránh bị chia nhỏ bởi Gemini
                        
                        // Tính scale factor để fit trong 768x768 (tránh bị chia nhỏ)
                        double scaleForMax = Math.Min(
                            (double)maxDimension / maxWidth,
                            (double)maxDimension / totalHeight
                        );
                        
                        // Đảm bảo không < 384x384 (giữ chất lượng tối thiểu)
                        double scaleForMin = Math.Max(
                            (double)minDimension / maxWidth,
                            (double)minDimension / totalHeight
                        );
                        
                        // Scale factor ban đầu để giảm dung lượng (target: <800KB)
                        double scaleForSize = 0.7; // Giảm 30% kích thước
                        
                        // Ước tính và điều chỉnh nếu vẫn quá lớn
                        long estimatedSize = (long)(maxWidth * scaleForSize * totalHeight * scaleForSize * 3 * 0.3); // ~30% của ảnh RGB sau nén
                        int maxRetries = 3;
                        int retry = 0;
                        
                        while (estimatedSize > 800 * 1024 && retry < maxRetries)
                        {
                            scaleForSize *= 0.85; // Giảm thêm 15% mỗi lần
                            estimatedSize = (long)(maxWidth * scaleForSize * totalHeight * scaleForSize * 3 * 0.3);
                            retry++;
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Ước tính size vẫn lớn ({estimatedSize / 1024}KB), giảm scale xuống {scaleForSize:F2}...");
                        }
                        
                        // Scale factor cuối cùng: đảm bảo trong khoảng min-max và target size
                        double finalScale = Math.Max(scaleForMin, Math.Min(scaleForMax, scaleForSize));
                        
                        // Tính kích thước cuối cùng
                        int targetWidth = (int)Math.Round(maxWidth * finalScale);
                        int targetHeight = (int)Math.Round(totalHeight * finalScale);
                        
                        // Đảm bảo không vượt quá giới hạn Gemini (768x768)
                        if (targetWidth > maxDimension || targetHeight > maxDimension)
                        {
                            double maxScale = Math.Min((double)maxDimension / targetWidth, (double)maxDimension / targetHeight);
                            targetWidth = (int)Math.Round(targetWidth * maxScale);
                            targetHeight = (int)Math.Round(targetHeight * maxScale);
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Điều chỉnh để không vượt quá {maxDimension}x{maxDimension}: {targetWidth}x{targetHeight}");
                        }
                        
                        // Đảm bảo không nhỏ hơn tối thiểu (384x384) - nhưng nếu ảnh gốc nhỏ thì giữ nguyên
                        if (targetWidth < minDimension && maxWidth >= minDimension)
                        {
                            targetWidth = minDimension;
                            targetHeight = (int)Math.Round(totalHeight * (double)minDimension / maxWidth);
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Điều chỉnh để không nhỏ hơn {minDimension}px width: {targetWidth}x{targetHeight}");
                        }
                        if (targetHeight < minDimension && totalHeight >= minDimension)
                        {
                            targetHeight = minDimension;
                            targetWidth = (int)Math.Round(maxWidth * (double)minDimension / totalHeight);
                            Console.WriteLine($"[AdLink] [CaptureFullPage] Điều chỉnh để không nhỏ hơn {minDimension}px height: {targetWidth}x{targetHeight}");
                        }
                        
                        // Giảm JPEG quality để giảm thêm dung lượng (60 thay vì 70-80)
                        int finalQuality = Math.Min(60, (int)Math.Clamp(jpegQuality, 1, 100));
                        
                        Console.WriteLine($"[AdLink] [CaptureFullPage] Resize từ {maxWidth}x{totalHeight} xuống {targetWidth}x{targetHeight} (scale={finalScale:F2}), quality={finalQuality}");
                        Console.WriteLine($"[AdLink] [CaptureFullPage] Kích thước cuối cùng: {targetWidth}x{targetHeight} (Gemini optimal: {minDimension}-{maxDimension}px, target: 258 tokens)");
                        
                        using (var resizedImage = mergedImage.Clone(ctx => ctx.Resize(new ResizeOptions
                        {
                            Mode = ResizeMode.Stretch,
                            Size = new Size(Math.Max(1, targetWidth), Math.Max(1, targetHeight))
                        })))
                        {
                            using (var outputMs = new MemoryStream())
                            {
                                var encoder = new JpegEncoder { Quality = finalQuality };
                                resizedImage.Save(outputMs, encoder);
                                var compressedBytes = outputMs.ToArray();
                                string base64Result = Convert.ToBase64String(compressedBytes);
                                
                                double sizeInKB = compressedBytes.Length / 1024.0;
                                Console.WriteLine($"[AdLink] [CaptureFullPage] SUCCESS: Đã ghép {images.Count} ảnh, kích thước sau nén: {compressedBytes.Length} bytes ({sizeInKB:F2}KB), base64: {base64Result.Length} ký tự");
                                
                                if (compressedBytes.Length > 800 * 1024)
                                {
                                    Console.WriteLine($"[AdLink] [CaptureFullPage] WARNING: Kích thước vẫn >800KB ({sizeInKB:F2}KB), có thể cần giảm quality hoặc scale thêm");
                                }
                                
                                return base64Result;
                            }
                        }
                    }
                }
                finally
                {
                    // Dispose tất cả images
                    foreach (var img in images)
                    {
                        try { img?.Dispose(); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [CaptureFullPage] EXCEPTION: Lỗi khi chụp full page screenshot: {ex.Message}");
                Console.WriteLine($"[AdLink] [CaptureFullPage] StackTrace: {ex.StackTrace}");
                ErrorWriter.WriteLog(LogPath, "CaptureFullPageScreenshotAsBase64", ex.ToString());
                
                // Đảm bảo quay lại trang gốc nếu có lỗi
                if (!string.IsNullOrWhiteSpace(originalUrl))
                {
                    try
                    {
                        TryNavigate(driver, originalUrl, TimeSpan.FromSeconds(30), out _);
                    }
                    catch { }
                }
                
                return "";
            }
        }
        
        /// <summary>
        /// Cuộn trang một cách mượt mà lên đầu trang để đảm bảo tất cả nội dung đã load (giống ScreenDesktop.cs)
        /// </summary>
        private static void SmoothScrollToTopForScreenshot(IWebDriver driver)
        {
            var js = (IJavaScriptExecutor)driver;
            try
            {
                long totalHeight = 0;
                try
                {
                    totalHeight = Convert.ToInt64(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
                }
                catch { }

                if (totalHeight <= 0) return;

                const int scrollSteps = 10;
                long stepSize = totalHeight / scrollSteps;
                
                for (int i = scrollSteps; i >= 0; i--)
                {
                    long scrollY = i * stepSize;
                    try
                    {
                        js.ExecuteScript($"window.scrollTo(0, {scrollY});");
                    }
                    catch { }
                    System.Threading.Thread.Sleep(150);
                }

                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                System.Threading.Thread.Sleep(300);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] [SmoothScrollToTop] Lỗi: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "SmoothScrollToTopForScreenshot", ex.ToString());
            }
        }
        
        /// <summary>
        /// Thử đóng các popup phổ biến trên trang đích trước khi chụp screenshot
        /// </summary>
        private static void CloseCommonPopups(IWebDriver driver)
        {
            try
            {
                var js = driver as IJavaScriptExecutor;
                bool foundPopup = false;
                var selectors = new[]
                {
                    "button[aria-label*='close']",
                    "button[aria-label*='Close']",
                    "button[aria-label*='Đóng']",
                    "div[role='button'][aria-label*='close']",
                    "div[role='button'][aria-label*='Close']",
                    "[data-dismiss='modal']",
                    ".modal-close, .modal__close, .close-button, .popup-close, .btn-close, .c-modal__close",
                    ".ant-modal-close, .ant-modal-close-x, .ant-modal-close-icon",
                    ".mfp-close, .fancybox-button--close, .fancybox-close-small",
                    ".pswp__button--close",
                    ".close, .Close"
                };

                foreach (var selector in selectors)
                {
                    IReadOnlyCollection<IWebElement> elements;
                    try
                    {
                        elements = driver.FindElements(By.CssSelector(selector));
                    }
                    catch
                    {
                        continue;
                    }

                    if (elements.Count == 0) continue;
                    foundPopup = true;

                    foreach (var element in elements)
                    {
                        try
                        {
                            if (!element.Displayed || !element.Enabled) continue;

                            try
                            {
                                element.Click();
                            }
                            catch
                            {
                                js?.ExecuteScript("arguments[0].click();", element);
                            }

                            Console.WriteLine($"[AdLink] Đã đóng popup bằng selector: {selector}");
                            System.Threading.Thread.Sleep(200);
                            break; // đủ đóng 1 popup theo selector
                        }
                        catch { }
                    }
                }

                // Nếu chưa thấy popup nào, thử kiểm tra overlay
                if (!foundPopup)
                {
                    IReadOnlyCollection<IWebElement>? overlays = null;
                    try
                    {
                        overlays = driver.FindElements(By.CssSelector(".modal-backdrop, .ant-modal-mask, .overlay, .backdrop, .ReactModal__Overlay"));
                    }
                    catch { }

                    if (overlays != null && overlays.Count > 0)
                    {
                        foundPopup = true;
                        foreach (var overlay in overlays)
                        {
                            try
                            {
                                if (!overlay.Displayed) continue;
                                js?.ExecuteScript("arguments[0].click();", overlay);
                                Console.WriteLine("[AdLink] Đã click overlay/backdrop để đóng popup");
                                break;
                            }
                            catch { }
                        }
                    }
                }

                if (!foundPopup)
                {
                    Console.WriteLine("[AdLink] Không phát hiện popup, bỏ qua bước đóng");
                    return;
                }

                // Gửi phím ESC như fallback
                try
                {
                    new Actions(driver).SendKeys(Keys.Escape).Perform();
                    Console.WriteLine("[AdLink] Gửi phím ESC để đóng popup");
                }
                catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] Lỗi khi đóng popup trang đích: {ex.Message}");
            }
        }
        
        private static bool TryClickIframeElement(IWebDriver driver, IWebElement iframeElement, int centerX, int centerY)
        {
            bool clicked = false;
            try
            {
                var actions = new Actions(driver);
                actions.MoveToElement(iframeElement, iframeElement.Size.Width / 2, iframeElement.Size.Height / 2)
                       .Click()
                       .Perform();
                clicked = true;
                Console.WriteLine("[AdLink] Đã click iframe bằng Actions tại center");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] Click iframe bằng Actions thất bại: {ex.Message}");
            }

            if (!clicked)
            {
                try
                {
                    var js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript(@"
                        var x = arguments[0];
                        var y = arguments[1];
                        var element = document.elementFromPoint(x, y);
                        if (element) {
                            var evt = new MouseEvent('click', {
                                view: window,
                                bubbles: true,
                                cancelable: true,
                                clientX: x,
                                clientY: y,
                                button: 0
                            });
                            element.dispatchEvent(evt);
                            return true;
                        }
                        return false;
                    ", centerX, centerY);
                    clicked = true;
                    Console.WriteLine("[AdLink] Đã click iframe bằng JavaScript (viewport)");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[AdLink] Click iframe bằng JavaScript thất bại: {ex.Message}");
                }
            }

            return clicked;
        }

        private static List<IWebElement> CollectBannerElements(IWebDriver driver)
        {
            var bannerElements = new List<IWebElement>();
            try
            {
                var iframes = driver.FindElements(By.TagName("iframe"));
                Console.WriteLine($"[AdLink] Tìm thấy {iframes.Count} iframe(s) trên trang");
                foreach (var iframe in iframes)
                {
                    try
                    {
                        var js = (IJavaScriptExecutor)driver;
                        var isVisible = (bool)js.ExecuteScript(@"
                                var rect = arguments[0].getBoundingClientRect();
                                return rect.width >= 120 && rect.height >= 30 && 
                                       rect.top >= 0 && rect.left >= 0 &&
                                       window.getComputedStyle(arguments[0]).display !== 'none';
                            ", iframe);

                        if (isVisible)
                        {
                            bannerElements.Add(iframe);
                        }
                    }
                    catch { }
                }

                var adSelectors = GetCommonAdSelectors();
                Console.WriteLine($"[AdLink] Bắt đầu tìm element quảng cáo với {adSelectors.Length} selector(s)");

                foreach (var selector in adSelectors)
                {
                    try
                    {
                        var adElements = driver.FindElements(By.CssSelector(selector));
                        Console.WriteLine($"[AdLink] Selector '{selector}': tìm thấy {adElements.Count} element(s)");

                        foreach (var element in adElements)
                        {
                            try
                            {
                                var js = (IJavaScriptExecutor)driver;
                                var isClickable = (bool)js.ExecuteScript(@"
                                        var el = arguments[0];
                                        var rect = el.getBoundingClientRect();
                                        if (rect.width < 120 || rect.height < 30) return false;
                                        if (rect.top < 0 || rect.left < 0) return false;
                                        var style = window.getComputedStyle(el);
                                        if (style.display === 'none' || style.visibility === 'hidden') return false;
                                        return true;
                                    ", element);

                                if (isClickable)
                                {
                                    bannerElements.Add(element);
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[AdLink] Lỗi khi xử lý selector '{selector}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] Lỗi khi thu thập banner element: {ex.Message}");
            }

            if (bannerElements.Count == 0)
            {
                Console.WriteLine("[AdLink] Không tìm thấy banner nào trong lần quét này");
            }
            else
            {
                Console.WriteLine($"[AdLink] Lần quét hiện tại thu được {bannerElements.Count} banner(s)");
            }

            return bannerElements;
        }

        private static string? TryGetBannerFingerprint(IWebDriver driver, IWebElement element)
        {
            try
            {
                var js = driver as IJavaScriptExecutor;
                var rectInfo = js?.ExecuteScript(@"
                        var rect = arguments[0].getBoundingClientRect();
                        return {
                            left: Math.round(rect.left),
                            top: Math.round(rect.top),
                            width: Math.round(rect.width),
                            height: Math.round(rect.height)
                        };
                    ", element) as System.Collections.Generic.Dictionary<string, object>;

                string rectKey = "rect:unknown";
                if (rectInfo != null &&
                    rectInfo.ContainsKey("left") &&
                    rectInfo.ContainsKey("top") &&
                    rectInfo.ContainsKey("width") &&
                    rectInfo.ContainsKey("height"))
                {
                    rectKey = $"{rectInfo["left"]}:{rectInfo["top"]}:{rectInfo["width"]}:{rectInfo["height"]}";
                }

                string representation = element.GetAttribute("outerHTML") ?? element.Text ?? element.TagName ?? "unknown";
                if (representation.Length > 200)
                {
                    representation = representation.Substring(0, 200);
                }

                return $"{rectKey}|{representation}";
            }
            catch (StaleElementReferenceException)
            {
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] TryGetBannerFingerprint gặp lỗi: {ex.Message}");
                return Guid.NewGuid().ToString();
            }
        }

        private static IWebElement? GetNextBannerElement(
            IWebDriver driver,
            IReadOnlyList<IWebElement> candidates,
            HashSet<string> processedFingerprints,
            out string? fingerprint,
            out int index)
        {
            fingerprint = null;
            index = -1;

            for (int i = 0; i < candidates.Count; i++)
            {
                var candidate = candidates[i];
                var candidateFingerprint = TryGetBannerFingerprint(driver, candidate);
                if (string.IsNullOrWhiteSpace(candidateFingerprint))
                {
                    continue;
                }

                if (processedFingerprints.Contains(candidateFingerprint))
                {
                    continue;
                }

                fingerprint = candidateFingerprint;
                index = i;
                return candidate;
            }

            return null;
        }

        private static bool TryGetNewWindowHandle(IWebDriver driver, string originalHandle, int originalCount, TimeSpan timeout, out string? newHandle)
        {
            newHandle = null;
            var sw = Stopwatch.StartNew();
            while (sw.Elapsed < timeout)
            {
                try
                {
                    var handles = driver.WindowHandles.ToList();
                    if (handles.Count > originalCount)
                    {
                        newHandle = handles.LastOrDefault(h => h != originalHandle);
                        return newHandle != null;
                    }
                }
                catch (WebDriverException ex)
                {
                    Console.WriteLine($"[AdLink] Lỗi khi đọc window handles: {ex.Message}, thử lại...");
                }
                System.Threading.Thread.Sleep(250);
            }
            return false;
        }

        private static void CloseAllNonOriginalTabs(IWebDriver driver, ref string originalHandle)
        {
            try
            {
                var handlesSnapshot = driver.WindowHandles.ToList();
                if (handlesSnapshot.Count == 0)
                {
                    Console.WriteLine("[AdLink] Không có window handle nào để đóng");
                    return;
                }

                if (!handlesSnapshot.Contains(originalHandle))
                {
                    originalHandle = handlesSnapshot[0];
                    Console.WriteLine("[AdLink] Tab gốc không còn, cập nhật originalHandle mới");
                }

                foreach (var handle in handlesSnapshot)
                {
                    if (handle == originalHandle) continue;

                    try
                    {
                        driver.SwitchTo().Window(handle);
                        driver.Close();
                        Console.WriteLine($"[AdLink] Đã đóng tab phụ: {handle}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[AdLink] Lỗi khi đóng tab {handle}: {ex.Message}");
                    }
                }

                try
                {
                    driver.SwitchTo().Window(originalHandle);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[AdLink] Lỗi khi switch về tab gốc: {ex.Message}");
                    var latestHandles = driver.WindowHandles;
                    if (latestHandles.Count > 0)
                    {
                        originalHandle = latestHandles[0];
                        driver.SwitchTo().Window(originalHandle);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AdLink] CloseAllNonOriginalTabs gặp lỗi: {ex.Message}");
            }
        }

        private static bool TryWaitForUrlChange(IWebDriver driver, string originalUrl, TimeSpan timeout, out string newUrl)
        {
            newUrl = originalUrl;
            var sw = Stopwatch.StartNew();
            while (sw.Elapsed < timeout)
            {
                try
                {
                    var currentUrl = driver.Url;
                    if (!string.Equals(currentUrl, originalUrl, StringComparison.OrdinalIgnoreCase))
                    {
                        newUrl = currentUrl;
                        return true;
                    }
                }
                catch (WebDriverException ex)
                {
                    Console.WriteLine($"[AdLink] Lỗi khi đọc URL: {ex.Message}, thử lại...");
                }
                System.Threading.Thread.Sleep(250);
            }
            return false;
        }

        private static bool WaitForPageLoad(IWebDriver driver, TimeSpan timeout)
        {
            var js = driver as IJavaScriptExecutor;
            var sw = Stopwatch.StartNew();
            while (sw.Elapsed < timeout)
            {
                try
                {
                    var state = js?.ExecuteScript("return document.readyState")?.ToString();
                    if (state == "complete" || state == "interactive")
                    {
                        return true;
                    }
                }
                catch (WebDriverException ex)
                {
                    Console.WriteLine($"[AdLink] Lỗi khi kiểm tra document.readyState: {ex.Message}, thử lại...");
                }
                System.Threading.Thread.Sleep(500);
            }
            return false;
        }
        
        /// <summary>
        /// Push link iframe vào RabbitMQ queue RabbitQueueAnalyzeSingleBanner
        /// 
        /// Cách thức hoạt động:
        /// 1. Kiểm tra queue đã được cấu hình
        /// 2. Tạo JSON payload: 
        ///    - Merge tất cả thông tin từ job params gốc (link_web, slice, quanlity_image, device, ...)
        ///    - Thêm "link_click_banner": "{link}"
        ///    - Thêm "screenshot_base64": "{base64_string}"
        /// 3. Đảm bảo kết nối RabbitMQ đã sẵn sàng
        /// 4. Push message lên queue với persistent = true
        /// 
        /// Tham số:
        /// - linkClick: Link click của iframe/quảng cáo
        /// - screenshotBase64: Base64 string của screenshot (từ CaptureFullPageScreenshot)
        /// - jobParamsSnapshot: Thông tin từ job gốc (RabbitQueue) để merge vào payload
        /// </summary>
        private static void PushIframeToQueue(string linkClick, string screenshotBase64, JObject? jobParamsSnapshot = null, string contactInfoJson = "{}")
        {
            Console.WriteLine($"[Iframe] ========== HÀM PushIframeToQueue ĐƯỢC GỌI ==========");
            Console.WriteLine($"[Iframe] Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            Console.WriteLine($"[Iframe] LinkClick: '{linkClick}'");
            Console.WriteLine($"[Iframe] ScreenshotBase64 Length: {(screenshotBase64?.Length ?? 0)}");
            Console.WriteLine($"[Iframe] ContactInfoJson Length: {(contactInfoJson?.Length ?? 0)}");
            Console.WriteLine($"[Iframe] JobParamsSnapshot IsNull: {jobParamsSnapshot == null}");
            
            try
            {
            
                // Gửi link tới queue screen link để job chụp ảnh từ link click xử lý
                var targetQueue = RabbitQueueScreenLink; // Queue name này dùng để nhận dữ liệu từ link click banner
                Console.WriteLine($"[Iframe] Chuẩn bị push queue. Queue='{targetQueue}', Link='{linkClick}', ScreenshotLen={(screenshotBase64?.Length ?? 0)}, ContactInfoLen={(contactInfoJson?.Length ?? 0)}");
                
                if (string.IsNullOrWhiteSpace(linkClick))
                {
                    Console.WriteLine("[Iframe] Link click rỗng, bỏ qua");
                    return;
                }
                
                if (string.IsNullOrWhiteSpace(targetQueue))
                {
                    Console.WriteLine("[Iframe] RabbitQueueScreenLink chưa được cấu hình, không thể push");
                    return;
                }
                
                if (string.IsNullOrWhiteSpace(screenshotBase64))
                {
                    Console.WriteLine("[Iframe] Screenshot rỗng, bỏ qua không push vào queue");
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
                
                // Thêm link_click_banner, screenshot_base64, contact_info_extracted và page_source (ghi đè nếu có trong job params)
                mergedPayload["link_click_banner"] = linkClick ?? "";
                mergedPayload["screenshot_base64"] = screenshotBase64 ?? "";
                mergedPayload["contact_info_extracted"] = contactInfoJson ?? "{}";
                
                var jsonPayload = mergedPayload.ToString(Newtonsoft.Json.Formatting.None);
                var body = Encoding.UTF8.GetBytes(jsonPayload);
                Console.WriteLine($"[Iframe] Payload chuẩn bị gửi có size={body.Length} bytes");
                
                // Debug: Log payload để kiểm tra
                try
                {
                    var payloadKeys = string.Join(", ", mergedPayload.Properties().Select(p => p.Name));
                    Console.WriteLine($"[Iframe] Debug - Payload keys: {payloadKeys}");
                }
                catch { }
                
                // Đảm bảo kết nối RabbitMQ sẵn sàng với queue name từ config
                EnsureIframePublisherReady(targetQueue);
                Console.WriteLine($"[Iframe] Kiểm tra channel RabbitMQ: {(iframeChannel != null ? (iframeChannel.IsOpen ? "Đang mở" : "Đang đóng") : "NULL")}");
                
                if (iframeChannel == null || !iframeChannel.IsOpen)
                {
                    Console.WriteLine("[Iframe] Không thể kết nối tới RabbitMQ");
                    return;
                }
                
                var props = iframeChannel.CreateBasicProperties();
                props.Persistent = true;
                
                // ========== LOG CHI TIẾT TRƯỚC KHI PUSH VÀO QUEUE ==========
                Console.WriteLine($"[Iframe] ========== BẮT ĐẦU PUSH VÀO QUEUE ==========");
                Console.WriteLine($"[Iframe] Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                Console.WriteLine($"[Iframe] Queue Name: '{targetQueue}'");
                Console.WriteLine($"[Iframe] Routing Key: '{targetQueue}'");
                Console.WriteLine($"[Iframe] Exchange: '' (default)");
                Console.WriteLine($"[Iframe] Body Size: {body.Length} bytes");
                Console.WriteLine($"[Iframe] Link Click: {linkClick}");
                Console.WriteLine($"[Iframe] Link Click Length: {linkClick?.Length ?? 0}");
                Console.WriteLine($"[Iframe] Screenshot Base64 Length: {(screenshotBase64?.Length ?? 0)}");
                Console.WriteLine($"[Iframe] Contact Info Length: {(contactInfoJson?.Length ?? 0)}");
                Console.WriteLine($"[Iframe] Persistent: {props.Persistent}");
                Console.WriteLine($"[Iframe] Channel IsOpen: {iframeChannel.IsOpen}");
                Console.WriteLine($"[Iframe] Channel IsNull: {iframeChannel == null}");
                Console.WriteLine($"[Iframe] Connection IsOpen: {(iframeConnection?.IsOpen ?? false)}");
                
                // Log một phần payload để debug (cắt ngắn nếu quá dài)
                try
                {
                    var payloadPreview = jsonPayload.Length > 500 
                        ? jsonPayload.Substring(0, 500) + "... [truncated]" 
                        : jsonPayload;
                    Console.WriteLine($"[Iframe] Payload Preview (first 500 chars): {payloadPreview}");
                }
                catch (Exception logEx)
                {
                    Console.WriteLine($"[Iframe] Không thể log payload preview: {logEx.Message}");
                }
                
                // Log các keys trong payload
                try
                {
                    var payloadKeys = string.Join(", ", mergedPayload.Properties().Select(p => p.Name));
                    Console.WriteLine($"[Iframe] Payload Keys: [{payloadKeys}]");
                }
                catch (Exception logEx)
                {
                    Console.WriteLine($"[Iframe] Không thể log payload keys: {logEx.Message}");
                }
                
                Console.WriteLine($"[Iframe] Sẵn sàng gọi BasicPublish...");
                
                iframeChannel.BasicPublish(exchange: "",
                                          routingKey: targetQueue,
                                          basicProperties: props,
                                          body: body);
                
                Console.WriteLine($"[Iframe] ========== ĐÃ PUSH THÀNH CÔNG VÀO QUEUE ==========");
                Console.WriteLine($"[Iframe] Đã push vào queue '{targetQueue}' - Link: {linkClick}, Size: {body.Length} bytes");
                Console.WriteLine($"[Iframe] Timestamp sau khi push: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Iframe] Lỗi push vào queue: {ex.Message}");
                Console.WriteLine($"[Iframe] Thông tin debug -> Link: {linkClick}, ScreenshotLen: {(screenshotBase64?.Length ?? 0)}");
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
                if (imageBytes == null || imageBytes.Length == 0)
                {
                    Console.WriteLine("[TryPublishAnalyze] Skipped: imageBytes is null or empty");
                    return;
                }
                
                if (string.IsNullOrWhiteSpace(RabbitQueueAnalyze))
                {
                    Console.WriteLine("[TryPublishAnalyze] Skipped: RabbitQueueAnalyze is not configured");
                    return;
                }
                
                EnsureAnalyzePublisherReady();
                
                // Fix: Kiểm tra cả null VÀ IsOpen
                if (analyzeChannel == null || !analyzeChannel.IsOpen)
                {
                    Console.WriteLine($"[TryPublishAnalyze] Skipped: analyzeChannel is null or closed (null={analyzeChannel == null}, IsOpen={analyzeChannel?.IsOpen ?? false})");
                    return;
                }

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
                Console.WriteLine($"[TryPublishAnalyze] Exception: {ex.Message}");
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
                    
                    // Thêm log thành công
                    Console.WriteLine($"[EnsureAnalyzePublisherReady] Successfully connected to RabbitMQ and declared queue: '{RabbitQueueAnalyze}'");
                }
                catch (Exception ex)
                {
                    // Thêm Console.WriteLine để dễ debug
                    Console.WriteLine($"[EnsureAnalyzePublisherReady] Error: {ex.Message}");
                    ErrorWriter.WriteLog(LogPath, "EnsureAnalyzePublisher", ex.ToString());
                    
                    // Đảm bảo channel = null khi fail
                    analyzeChannel?.Dispose();
                    analyzeChannel = null;
                    analyzeConnection?.Dispose();
                    analyzeConnection = null;
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
            int fallbackAttempts = 0;
            var deadline = DateTime.UtcNow + maxWait;

            try
            {
                lastHeight = Convert.ToInt64(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
            }
            catch { }

            ForceEnableBodyScroll(driver);

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

                    if (stableRounds >= 2 && fallbackAttempts < 4)
                    {
                        Console.WriteLine($"[ScrollFallback] Trang không tăng chiều cao sau {stableRounds} vòng. Thử fallback #{fallbackAttempts + 1}...");

                        ForceEnableBodyScroll(driver);
                        bool containerScrolled = TryScrollScrollableContainers(driver);
                        if (!containerScrolled)
                        {
                            Console.WriteLine("[ScrollFallback] Không tìm thấy container scrollable hoặc không scroll được. Thử gửi wheel event/PageDown...");
                            TryDispatchWheelFallback(driver);
                        }
                        else
                        {
                            Console.WriteLine("[ScrollFallback] Đã scroll trong container nội bộ.");
                        }

                        stableRounds = 0;
                        fallbackAttempts++;
                        continue;
                    }
                }
                else
                {
                    stableRounds = 0;
                    lastHeight = newHeight;
                }

                if (stableRounds >= 3) break;
            }

            if (fallbackAttempts > 0)
            {
                Console.WriteLine($"[ScrollFallback] Tổng số lần fallback đã thực hiện: {fallbackAttempts}");
            }

            // Một lần chờ ngắn để JS còn lại hoàn tất
            WaitForReadyState(driver, TimeSpan.FromSeconds(2));
        }

        private static void ForceEnableBodyScroll(IWebDriver driver)
        {
            var js = (IJavaScriptExecutor)driver;
            try
            {
                js.ExecuteScript(@"
                    try {
                        document.body.style.overflow = 'auto';
                        document.body.style.position = 'static';
                    } catch(e) {}
                    try {
                        document.documentElement.style.overflow = 'auto';
                        document.documentElement.style.position = 'static';
                    } catch(e) {}
                ");
            }
            catch { }
        }

        private static bool TryScrollScrollableContainers(IWebDriver driver)
        {
            var js = (IJavaScriptExecutor)driver;
            try
            {
                var result = js.ExecuteScript(@"
                    let scrolled = false;
                    const candidates = Array.from(document.querySelectorAll('*'));
                    for (const el of candidates) {
                        try {
                            const style = window.getComputedStyle(el);
                            if (!style) continue;
                            const overflowY = style.overflowY || style.overflow;
                            const canScroll = (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay');
                            if (!canScroll) continue;
                            const maxScroll = el.scrollHeight - el.clientHeight;
                            if (maxScroll > 20) {
                                const before = el.scrollTop;
                                el.scrollTop = el.scrollHeight;
                                if (Math.abs(el.scrollTop - before) > 5) {
                                    scrolled = true;
                                }
                            }
                        } catch {}
                    }
                    return scrolled;
                ");
                return result is bool b && b;
            }
            catch
            {
                return false;
            }
        }

        private static void TryDispatchWheelFallback(IWebDriver driver, int deltaY = 800)
        {
            bool success = false;
            try
            {
                new Actions(driver).ScrollByAmount(0, deltaY).Perform();
                success = true;
                Console.WriteLine("[ScrollFallback] Đã scroll bằng Actions.ScrollByAmount");
            }
            catch { }

            if (!success)
            {
                try
                {
                    new Actions(driver).SendKeys(Keys.PageDown).Perform();
                    success = true;
                    Console.WriteLine("[ScrollFallback] Đã scroll bằng phím PageDown");
                }
                catch { }
            }

            if (!success)
            {
                try
                {
                    var js = (IJavaScriptExecutor)driver;
                    int centerX = 500;
                    int centerY = 500;
                    try
                    {
                        var dict = js.ExecuteScript("return { x: Math.floor(window.innerWidth/2), y: Math.floor(window.innerHeight/2) };") as Dictionary<string, object>;
                        if (dict != null)
                        {
                            centerX = Convert.ToInt32(dict["x"]);
                            centerY = Convert.ToInt32(dict["y"]);
                        }
                    }
                    catch { }

                    var chrome = driver as ChromeDriver;
                    if (chrome != null)
                    {
                        chrome.ExecuteCdpCommand("Input.dispatchMouseWheelEvent", new Dictionary<string, object>
                        {
                            { "type", "mouseWheel" },
                            { "x", centerX },
                            { "y", centerY },
                            { "deltaX", 0 },
                            { "deltaY", deltaY }
                        });
                        Console.WriteLine("[ScrollFallback] Đã gửi Input.dispatchMouseWheelEvent qua CDP");
                        success = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ScrollFallback] Lỗi khi gửi wheel event qua CDP: {ex.Message}");
                }
            }

            if (!success)
            {
                Console.WriteLine("[ScrollFallback] Tất cả phương án scroll fallback đều thất bại");
            }
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