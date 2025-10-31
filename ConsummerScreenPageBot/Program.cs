using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Win32;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using System.Threading;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Net;

namespace ConsummerScreenPageBot
{
    class Program
    {
        private static string startupPath = AppDomain.CurrentDomain.BaseDirectory.Replace(@"\bin\Debug\net8.0", @"\");
        public static string rabbit_queue_name = ConfigurationManager.AppSettings["RabbitQueue"] ?? "";
        public static string RabbitQueueAnalyze = ConfigurationManager.AppSettings["RabbitQueueAnalyze"] ?? "";
        
        public static string rabbit_host = ConfigurationManager.AppSettings["RabbitHost"] ?? "";
        public static string rabbit_vhost = ConfigurationManager.AppSettings["RabbitVHost"] ?? "";
        public static int rabbit_port = Convert.ToInt32(ConfigurationManager.AppSettings["RabbitPort"] ?? "5672");
        public static string rabbit_username = ConfigurationManager.AppSettings["RabbitUserName"] ?? "";
        public static string rabbit_password = ConfigurationManager.AppSettings["RabbitPassword"] ?? "";    
        public static string rabbit_use_ssl = ConfigurationManager.AppSettings["RabbitUseSSL"] ?? "0";
        public static string LogPath = ConfigurationManager.AppSettings["PATHLOG"] ?? "logs";
        public static string is_headless = ConfigurationManager.AppSettings["is_headless"] ?? "0";
        public static string websites_config = ConfigurationManager.AppSettings["Websites"] ?? "";

        // Publisher for analyze queue
        private static readonly object analyzePubLock = new object();
        private static IConnection? analyzeConnection;
        private static IModel? analyzeChannel;

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
                }
                else
                {
                    // Dòng bổ sung, cảnh báo khi Chrome không tìm thấy
                    Console.WriteLine("Warning: Chrome binary NOT found. Please check Chrome is installed or adjust the path.");
                }

                // Fix 2: Cấu hình chống crash headless/windowless môi trường server/docker
                if (is_headless == "1")
                {
                    chrome_option.AddArgument("--headless=new");
                    chrome_option.AddArgument("--disable-gpu");
                    chrome_option.AddArgument("--window-size=1920,1080");
                    chrome_option.AddArgument("--disable-software-rasterizer");
                    chrome_option.AddArgument("--no-sandbox");
                    chrome_option.AddArgument("--disable-dev-shm-usage");
                }
                else
                {
                    chrome_option.AddArgument("--start-maximized"); // set full man hinh  
                }

                // Use isolated user profile (giữ nguyên)
                var userDataDir = Path.Combine(startupPath, "chrome-profile");
                if (!Directory.Exists(userDataDir)) Directory.CreateDirectory(userDataDir);
                chrome_option.AddArgument($"--user-data-dir={userDataDir}");

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

             
                // SE READY...
                // Tự động đồng bộ phiên bản ChromeDriver với bản Chrome đang cài trên máy.
                // Cơ chế: WebDriverManager sẽ kiểm tra version Chrome hiện tại, nếu thiếu hoặc lệch phiên bản
                // thì tự tải về và trỏ ChromeDriver tương ứng. Nhờ vậy khi Chrome nâng cấp, driver cũng luôn khớp.
                try
                {
                    // Force WebDriverManager to download a ChromeDriver matching the installed Chrome version
                    var installedChromeVersion = string.Empty;
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(chrome_option.BinaryLocation) && File.Exists(chrome_option.BinaryLocation))
                        {
                            var fv = FileVersionInfo.GetVersionInfo(chrome_option.BinaryLocation);
                            installedChromeVersion = fv.FileVersion ?? string.Empty;
                        }
                    }
                    catch { }

                    if (!string.IsNullOrWhiteSpace(installedChromeVersion))
                    {
                        Console.WriteLine($"Detected Chrome version: {installedChromeVersion}");
                        // Many WebDriverManager implementations accept a version prefix (e.g., 141) to resolve the latest patch of that major
                        var major = installedChromeVersion.Split('.')?.FirstOrDefault();
                        var versionHint = string.IsNullOrWhiteSpace(major) ? installedChromeVersion : major;
                        new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig(), versionHint);
                        Console.WriteLine($"WebDriverManager: ensured ChromeDriver for Chrome {versionHint}.");
                    }
                    else
                    {
                        // Fallback to default resolution if we could not read Chrome's version
                        new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
                        Console.WriteLine("WebDriverManager: ensured ChromeDriver (no explicit version hint).");
                    }
                }
                catch (Exception wdmEx)
                {
                    Console.WriteLine("WebDriverManager failed: " + wdmEx.Message);
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
                    using (var browers = new ChromeDriver(service, chrome_option, TimeSpan.FromMinutes(3)))
                    {  
                         // DNS & TCP reachability diagnostics before creating RabbitMQ connection
                        // try
                        // {
                        //     Console.WriteLine("Resolving host: " + rabbit_host);
                        //     var addresses = Dns.GetHostAddresses(rabbit_host);
                        //     Console.WriteLine("Resolved IPs: " + string.Join(", ", addresses.Select(a => a.ToString())));

                        //     using (var tcp = new TcpClient())
                        //     {
                        //         var connectTask = tcp.ConnectAsync(rabbit_host, rabbit_port);
                        //         if (!connectTask.Wait(TimeSpan.FromSeconds(5)))
                        //         {
                        //             Console.WriteLine($"TCP precheck: timeout connecting to {rabbit_host}:{rabbit_port} (5s)");
                        //         }
                        //         else if (!tcp.Connected)
                        //         {
                        //             Console.WriteLine($"TCP precheck: failed to connect to {rabbit_host}:{rabbit_port}");
                        //         }
                        //         else
                        //         {
                        //             Console.WriteLine($"TCP precheck: connected to {rabbit_host}:{rabbit_port}");
                        //         }
                        //     }
                        // }
                        // catch (Exception preEx)
                        // {
                        //     Console.WriteLine("TCP/DNS precheck error: " + preEx.Message);
                        // }
                        #region WAITING QUEUE
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
                                    siteUrl = jobj["link_web"]?.ToString() ?? "";
                                    segment_page = jobj["slice"] != null ? jobj["slice"].ToObject<int>() : 10;
                                    
                                    try
                                    {                                        
                                        Console.WriteLine("Navigate: " + siteUrl);
                                        if (!TryNavigate(browers, siteUrl, TimeSpan.FromSeconds(60), out var failReason))
                                        {
                                            Console.WriteLine($"Navigation failed: {failReason}");
                                            ErrorWriter.WriteLog(LogPath, "NavigateFail", $"{siteUrl} => {failReason}");
                                            
                                        }
                                        ProcessWebsite(browers, siteUrl, segment_page);
                                    }
                                    catch (Exception siteEx)
                                    {
                                        Console.WriteLine($"Error processing {siteUrl}: {siteEx.Message}");
                                        ErrorWriter.WriteLog(LogPath, "ProcessWebsite", $"{siteUrl} => {siteEx}");
                                    }
                                    

                                    channel.BasicAck(deliveryTag: ea.DeliveryTag, multiple: false);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("error queue: " + ex.ToString());
                                    ErrorWriter.WriteLog(LogPath, "QueueError", ex.ToString());
                                }
                            };

                            channel.BasicConsume(queue: rabbit_queue_name, autoAck: false, consumer: consumer);

                            Console.ReadLine();

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                            throw;
                        }
                    }
                    
                        #endregion

                        browers.Dispose();
                    }
                }
                catch (WebDriverException wdEx)
                {
                    Console.WriteLine("WebDriverException: " + wdEx.Message);
                    // Main required edit: Specific message for this error code
                    if (wdEx.Message != null && wdEx.Message.Contains("session not created: Chrome instance exited.", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("session not created: Chrome instance exited. Examine ChromeDriver verbose log to determine the cause. (SessionNotCreated)");
                        Console.WriteLine($"See verbose ChromeDriver log at: {Path.Combine(LogPath, "chromedriver.log")}");
                    }
                    else
                    {
                        Console.WriteLine("Possible cause: Chrome or ChromeDriver version mismatch, Chrome not installed, or missing dependencies on server (libgbm, libX11, ...).");
                    }

                    ErrorWriter.WriteLog(LogPath, "SessionNotCreated", wdEx.ToString());

                    if (!string.IsNullOrEmpty(chrome_option.BinaryLocation))
                    {
                        Console.WriteLine("Using detected Chrome binary: " + chrome_option.BinaryLocation);
                    }
                    else
                    {
                        Console.WriteLine("No Chrome binary detected! Try specifying an absolute path or installing Chrome browser.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.ToString());
                    ErrorWriter.WriteLog(LogPath, "Handle Error", ex.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error (outer): " + ex.ToString());
                ErrorWriter.WriteLog(LogPath, "Handle Error(outer)", ex.ToString());
            }
        }

        private static void ProcessWebsite(IWebDriver driver, string url,int segment_page = 10)
        {
            string host;
            try
            {
                host = new Uri(url).Host.ToLowerInvariant();
            }
            catch
            {
                host = url.ToLowerInvariant();
            }

            // Xóa toàn bộ ảnh cũ trong thư mục của site trước khi chụp lại
            ClearHostScreenshots(host);

            // Cuộn xuống cuối trang để kích hoạt lazy-load (quảng cáo/thành phần chậm) trước khi dò tìm
            ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(12));
            TryProbeAdCandidates(driver, TimeSpan.FromSeconds(3));

            switch (host)
            {
                case var h when h.Contains("vnexpress.net"):
                    HandleVnExpress(driver);
                    break;
                case var h when h.Contains("thanhnien.vn"):
                    HandleThanhNien(driver);
                    break;
                default:
                    Console.WriteLine($"No specific handler for host: {host}. Using generic banner capture.");
                    CaptureGenericBanners(driver, host);
                    break;
            }

            // Chụp thêm ảnh full page để đảm bảo lưu toàn bộ quảng cáo xuất hiện trên trang
            //CaptureFullPageScreenshot(driver, host);

            // Chụp segment chia đều toàn bộ chiều dài trang
            CaptureSegmentScreenshots(driver, host, segment_page);
        }

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

        // Quay lên đầu trang, kích lazy-load và chờ banner/iframe hiển thị gần đầu trang
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

        // Xóa tất cả file ảnh trong thư mục screenshots/<hostLabel>
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
            }
        }

        // Chụp full-page bằng cách đặt kích thước viewport theo kích thước tài liệu qua Chrome DevTools Protocol
        private static void CaptureFullPageScreenshot(IWebDriver driver, string hostLabel)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                int width = 1920;
                int height = 1080;
                try
                {
                    width = Convert.ToInt32(js.ExecuteScript("return Math.max(document.documentElement.clientWidth, window.innerWidth || 0);"));
                    height = Convert.ToInt32(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight, window.innerHeight || 0);"));
                    // Giới hạn tối đa để tránh lỗi bộ nhớ trên các trang cực dài
                    height = Math.Min(height, 30000);
                }
                catch { }

                // Cuộn hết và chờ ngắn để lazy-load hoàn tất
                ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(5));
                TryProbeAdCandidates(driver, TimeSpan.FromSeconds(2));

                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                // Đặt viewport theo kích thước tài liệu để chụp full
                var chrome = driver as ChromeDriver;
                if (chrome != null)
                {
                    var metrics = new Dictionary<string, object>
                    {
                        { "mobile", false },
                        { "width", width },
                        { "height", height },
                        { "deviceScaleFactor", 1 },
                        { "scale", 1 }
                    };
                    try { chrome.ExecuteCdpCommand("Emulation.setDeviceMetricsOverride", metrics); } catch { }
                }

                var shot = ((ITakesScreenshot)driver).GetScreenshot();
                var fileName = $"full_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.png";
                var savePath = Path.Combine(shotsDir, fileName);
                shot.SaveAsFile(savePath);
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "FullPageScreenshot", ex.ToString());
            }
        }

        // Chụp N phần bằng cách chia tổng chiều cao trang thành N đoạn bằng nhau
        private static void CaptureSegmentScreenshots(IWebDriver driver, string hostLabel, int segmentCount)
        {
            try
            {
                var js = (IJavaScriptExecutor)driver;
                int pageWidth = 1920;
                int totalHeight = 3000;
                try
                {
                    pageWidth = Convert.ToInt32(js.ExecuteScript("return Math.max(document.documentElement.clientWidth, window.innerWidth || 0);"));
                    totalHeight = Convert.ToInt32(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
                }
                catch { }

                if (totalHeight <= 0) totalHeight = 3000;
                if (segmentCount <= 0) segmentCount = 3;
                int segmentHeight = (int)Math.Ceiling(totalHeight / (double)segmentCount);
                // Giới hạn để tránh set quá lớn gây out-of-memory với DevTools
                segmentHeight = Math.Min(segmentHeight, 10000);

                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                // Phát hiện vị trí (top/bottom) của các khối quảng cáo để tránh cắt ngang khi chia đoạn
                var adRects = TryDetectAdRects(driver);

                var chrome = driver as ChromeDriver;
                if (chrome != null)
                {
                    // Cố định chiều rộng, thay đổi chiều cao theo segment
                    var baseMetrics = new Dictionary<string, object>
                    {
                        { "mobile", false },
                        { "width", pageWidth },
                        { "deviceScaleFactor", 1 },
                        { "scale", 1 }
                    };

                    int lastStart = 0;
                    for (int i = 0; i < segmentCount; i++)
                    {
                        int y = i * segmentHeight;
                        int currentHeight = Math.Min(segmentHeight, Math.Max(1, totalHeight - y));

                        // Set viewport theo chiều cao của segment hiện tại
                        var metrics = new Dictionary<string, object>(baseMetrics)
                        {
                            ["height"] = currentHeight
                        };
                        try { chrome.ExecuteCdpCommand("Emulation.setDeviceMetricsOverride", metrics); } catch { }

                        // Tìm toạ độ cuộn an toàn để tránh cắt ngang các khối quảng cáo tại biên trên/dưới
                        y = AdjustSliceStartToAvoidAds(y, currentHeight, totalHeight, adRects, lastStart);

                        // Cuộn đến offset tương ứng (đã điều chỉnh an toàn)
                        try { js.ExecuteScript("window.scrollTo(0, arguments[0]);", y); } catch { }
                        Thread.Sleep(350);

                        // Chụp ảnh viewport (chính là 1/3 trang)
                        var shot = ((ITakesScreenshot)driver).GetScreenshot();
                        var fileName = $"split{i+1}_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.png";
                        var savePath = Path.Combine(shotsDir, fileName);
                        shot.SaveAsFile(savePath);

                        // Gửi ảnh segment sang queue phân tích dưới dạng base64
                        TryPublishAnalyze(shot.AsByteArray);

                        // Đảm bảo tiến triển tăng dần, thêm một chồng lấn nhỏ để không bỏ sót nội dung
                        lastStart = Math.Max(lastStart, y + Math.Max(1, currentHeight - 20));
                    }
                }
                else
                {
                    // Fallback: nếu không phải ChromeDriver (CDP), chụp 1 ảnh/scroll (ít chính xác hơn)
                    for (int i = 0; i < segmentCount; i++)
                    {
                        int y = i * segmentHeight;
                        try { js.ExecuteScript("window.scrollTo(0, arguments[0]);", y); } catch { }
                        Thread.Sleep(350);
                        var shot = ((ITakesScreenshot)driver).GetScreenshot();
                        var fileName = $"split{i+1}_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.png";
                        var savePath = Path.Combine(startupPath, "screenshots", hostLabel, fileName);
                        Directory.CreateDirectory(Path.GetDirectoryName(savePath)!);
                        shot.SaveAsFile(savePath);

                        // Gửi ảnh segment sang queue phân tích dưới dạng base64
                        TryPublishAnalyze(shot.AsByteArray);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "SegmentScreenshots", ex.ToString());
            }
        }

        private static void TryPublishAnalyze(byte[] imageBytes)
        {
            try
            {
                if (imageBytes == null || imageBytes.Length == 0) return;
                if (string.IsNullOrWhiteSpace(RabbitQueueAnalyze)) return;
                EnsureAnalyzePublisherReady();
                if (analyzeChannel == null) return;

                var base64 = Convert.ToBase64String(imageBytes);
                var payload = $"{{ \"screenshot_base64\": \"{base64}\" }}";
                var body = Encoding.UTF8.GetBytes(payload);

                var props = analyzeChannel.CreateBasicProperties();
                props.Persistent = true;

                analyzeChannel.BasicPublish(exchange: "",
                                            routingKey: RabbitQueueAnalyze,
                                            basicProperties: props,
                                            body: body);
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "PublishAnalyze", ex.ToString());
            }
        }

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

        // Lấy danh sách các rect (top/bottom theo toạ độ tài liệu) của các phần tử có khả năng là quảng cáo
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

        // Hợp nhất các khoảng [top,bottom] chồng lấn
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

        // Điều chỉnh toạ độ bắt đầu của lát cắt sao cho biên trên/dưới không nằm giữa một khối quảng cáo
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

        private static void HandleThanhNien(IWebDriver driver)
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
            AdCapture.CaptureBySelectors(driver, selectors, "thanhnien.vn", startupPath, LogPath);
            AdCapture.CaptureAdIframes(driver, "thanhnien.vn", startupPath, LogPath);
        }

        private static void CaptureGenericBanners(IWebDriver driver, string hostLabel)
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
            AdCapture.CaptureBySelectors(driver, selectors, hostLabel, startupPath, LogPath);
            // Chụp trực tiếp các iframe
            AdCapture.CaptureAdIframes(driver, hostLabel, startupPath, LogPath);
        }

        // Wrapper giữ tương thích cũ, ủy thác sang Utils/AdCapture
        private static void CaptureBySelectors(IWebDriver driver, string[] selectors, string hostLabel)
        {
            AdCapture.CaptureBySelectors(driver, selectors, hostLabel, startupPath, LogPath);
        }

        // (moved to Utils/AdCapture.cs)

        /// <summary>
        /// Lưu lại ảnh chụp màn hình của một phần tử (element), cắt từ screenshot toàn trang, vào thư mục theo hostLabel.
        /// Nếu không cắt được (tọa độ vượt ngoài ảnh), sẽ lưu toàn bộ screenshot thay thế.
        /// </summary>
        // (moved to Utils/AdCapture.cs)

        // Tạo nhãn tên phần tử từ id hoặc class để đặt tên file
        // (moved to Utils/AdCapture.cs)

        // (moved to Utils/AdCapture.cs)

        // Generic: quay lên đầu trang, kích lazy-load và chờ banner/iframe hiện ở gần đầu trang (áp dụng rộng rãi)
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

        private static string FindChromeBinaryPath()
        {
            try
            {
                // Common install locations
                var candidates = new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Google\\Chrome\\Application\\chrome.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Google\\Chrome\\Application\\chrome.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google\\Chrome\\Application\\chrome.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Chromium\\Application\\chrome.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Chromium\\Application\\chrome.exe")
                };
                foreach (var p in candidates)
                {
                    if (File.Exists(p)) return p;
                }

                // Registry lookup
                string[] hives = { "HKEY_LOCAL_MACHINE", "HKEY_CURRENT_USER" };
                foreach (var hive in hives)
                {
                    var path = Registry.GetValue($"{hive}\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe", "", null) as string;
                    if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) return path;
                }
            }
            catch { }
            return string.Empty;
        }

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

        // Bộ selector chung áp dụng cho nhiều trang tin để bắt hầu hết container quảng cáo phổ biến
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

        // Cuộn xuống cuối trang theo từng bước để kích hoạt lazy-load; dừng khi chiều cao trang không tăng thêm vài lần liên tiếp
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

        // Thăm dò nhẹ xem các phần tử có khả năng là quảng cáo đã xuất hiện chưa (không bắt buộc)
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

    // (moved to Utils/ErrorWriter.cs)
}