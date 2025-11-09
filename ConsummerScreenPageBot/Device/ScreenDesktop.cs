using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;
using ConsummerScreenPageBot.Utils;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace ConsummerScreenPageBot.Device
{
    /// <summary>
    /// Class xử lý chụp màn hình cho Desktop device (PC)
    /// Viewport: 1920x1080, mobile: false
    /// </summary>
    public static class ScreenDesktop
    {
        private static string startupPath = AppDomain.CurrentDomain.BaseDirectory.Replace(@"\bin\Debug\net8.0", @"\");
        private static string LogPath = ConfigurationManager.AppSettings["PATHLOG"] ?? "logs";

        /// <summary>
        /// Xử lý website để chụp ảnh quảng cáo và nội dung trên Desktop
        /// 
        /// Cách thức hoạt động:
        /// 1. Lấy hostname từ URL để tạo thư mục lưu ảnh tương ứng
        /// 2. Xóa toàn bộ ảnh cũ trong thư mục screenshots/<hostname> để làm sạch
        /// 3. Cuộn trang xuống cuối để kích hoạt lazy-loading (quảng cáo, hình ảnh load chậm)
        ///    - Đợi cho đến khi chiều cao trang không tăng thêm (đã load hết nội dung)
        /// 4. Thăm dò các phần tử có thể là quảng cáo (banner, ads container)
        /// 5. Chụp ảnh theo segment: chia trang thành N phần bằng nhau và chụp từng phần
        ///    - Mỗi ảnh segment sẽ được nén với chất lượng jpegQuality
        ///    - Tự động gửi mỗi ảnh lên queue analyze để AI xử lý
        /// 
        /// Tham số:
        /// - driver: WebDriver đã điều hướng tới trang web
        /// - url: URL của trang web cần chụp
        /// - segment_page: Số lượng segment để chia trang (mặc định 10)
        /// - jpegQuality: Chất lượng ảnh JPEG từ 1-100 (mặc định 80)
        /// </summary>
        public static void ProcessWebsite(IWebDriver driver, string url, int segment_page = 10, long jpegQuality = 80)
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
            ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(15));
            TryProbeAdCandidates(driver, TimeSpan.FromSeconds(5));

            // Chụp segment chia đều toàn bộ chiều dài trang
            CaptureSegmentScreenshots(driver, host, segment_page, jpegQuality);

            // Sau khi chụp segment xong, xử lý tất cả iframe và push vào queue
               // Console.WriteLine("[Desktop] Bắt đầu xử lý iframe sau khi chụp segment...");
                //Program.ProcessIframesAndPushToQueue(driver, host);
        }

        /// <summary>
        /// Chụp ảnh trang web bằng cách chia thành nhiều segment (phần) và chụp từng phần - Desktop
        /// 
        /// Viewport Desktop: 1920x1080, mobile: false
        /// 
        /// Cách thức hoạt động:
        /// 1. Lấy kích thước thực tế của trang (width x height)
        ///    - Giới hạn chiều cao tối đa 30000px
        /// 2. Cuộn xuống cuối trang và đợi quảng cáo load hoàn toàn
        /// 3. Chụp ảnh toàn trang một lần bằng CDP:
        ///    - Đặt viewport theo kích thước tài liệu (mobile: false)
        ///    - Chụp ảnh full page bằng Page.captureScreenshot
        ///    - Nếu CDP fail => fallback về ITakesScreenshot
        /// 4. Load ảnh full page vào ImageSharp để xử lý
        /// 5. Chia ảnh thành N segment bằng nhau (slice)
        /// 6. Reset viewport về kích thước ban đầu
        /// </summary>
        private static void CaptureSegmentScreenshots(IWebDriver driver, string hostLabel, int segmentCount, long jpegQuality)
        {
            try
            {
                try { Console.WriteLine($"[Desktop Segment] Start capture host={hostLabel}, segments={segmentCount}, quality={jpegQuality}"); } catch { }
                if (segmentCount <= 0) segmentCount = 3;

                var js = (IJavaScriptExecutor)driver;
                
                // Bước 1: Scroll xuống cuối trang để kích hoạt lazy-load
                Console.WriteLine("[Desktop] Scroll xuống cuối trang để kích hoạt lazy-load...");
                ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(15));
                
                // Bước 2: Scroll ngược lại lên đầu trang một cách mượt mà để đảm bảo tất cả nội dung đã load
                Console.WriteLine("[Desktop] Scroll ngược lại lên đầu trang...");
                SmoothScrollToTop(driver);
                
                // Bước 3: Delay để đảm bảo tất cả dữ liệu (ảnh, quảng cáo, lazy content) đã load đầy đủ
                Console.WriteLine("[Desktop] Đợi nội dung load đầy đủ...");
                Thread.Sleep(2000);
                
                // Bước 4: Đợi quảng cáo load hoàn toàn
                WaitForAdsLoaded(driver, TimeSpan.FromSeconds(12));
                
                // Bước 5: Delay thêm để đảm bảo tất cả banner quảng cáo đã render hoàn toàn
                Thread.Sleep(1500);
                
                Console.WriteLine("[Desktop] Bắt đầu chụp màn hình...");
                
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

                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel, "desktop");
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                // 1) Chụp full page một lần bằng CDP; fallback sang ITakesScreenshot nếu cần
                byte[] fullShotBytes = Array.Empty<byte>();
                var chrome = driver as ChromeDriver;
                bool fullOk = false;
                if (chrome != null)
                {
                    var metrics = new Dictionary<string, object>
                    {
                        { "mobile", false }, // Desktop
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
                            var fileName = $"desktop_split{i+1}_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.jpg";
                            var savePath = Path.Combine(shotsDir, fileName);
                            AdCapture.SaveImageCompressed(seg, savePath, 1.0, jpegQuality);

                            try
                            {
                                var compressedBytes = File.ReadAllBytes(savePath);
                                try { Console.WriteLine($"[Desktop Segment] Saved {Path.GetFileName(savePath)} ({compressedBytes.Length} bytes)"); } catch { }
                                // Truyền width và height của segment
                                Program.TryPublishAnalyze(compressedBytes, seg.Width, seg.Height);
                            }
                            catch { }
                        }
                    }
                }
                try { Console.WriteLine($"[Desktop Segment] Done host={hostLabel}"); } catch { }
                
                
            }
            catch (Exception ex)
            {
                try { Console.WriteLine($"[Desktop Segment] ERROR {hostLabel}: {ex.Message}"); } catch { }
                ErrorWriter.WriteLog(LogPath, "DesktopSegmentScreenshots", ex.ToString());
                TelegramService.PushLogToTelegram($"Desktop SegmentScreenshots Error - {hostLabel}", ex);
            }
        }

        /// <summary>
        /// Xóa tất cả file ảnh cũ trong thư mục screenshots của host cụ thể - Desktop
        /// </summary>
        private static void ClearHostScreenshots(string hostLabel)
        {
            try
            {
                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel, "desktop");
                if (!Directory.Exists(shotsDir)) return;

                var files = Directory.GetFiles(shotsDir);
                foreach (var f in files)
                {
                    try { File.Delete(f); }
                    catch (Exception ex)
                    {
                        ErrorWriter.WriteLog(LogPath, "ClearScreenshots.Delete", $"Desktop {hostLabel} => {f} => {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "ClearScreenshots", $"Desktop {hostLabel} => {ex}");
                TelegramService.PushLogToTelegram($"Desktop ClearScreenshots Error - {hostLabel}", ex);
            }
        }

        /// <summary>
        /// Cuộn trang một cách mượt mà lên đầu trang để đảm bảo tất cả nội dung đã load
        /// </summary>
        private static void SmoothScrollToTop(IWebDriver driver)
        {
            var js = (IJavaScriptExecutor)driver;
            try
            {
                // Lấy chiều cao trang
                long totalHeight = 0;
                try
                {
                    totalHeight = Convert.ToInt64(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight) || 0;"));
                }
                catch { }

                if (totalHeight <= 0) return;

                // Scroll mượt mà lên đầu trang bằng cách chia thành nhiều bước nhỏ
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
                    Thread.Sleep(150); // Delay nhỏ giữa mỗi bước scroll
                }

                // Đảm bảo đã về đầu trang
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                Thread.Sleep(300); // Delay cuối cùng
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "SmoothScrollToTop", ex.ToString());
            }
        }

        /// <summary>
        /// Cuộn trang xuống cuối để kích hoạt lazy-loading và đảm bảo nội dung đã load hết
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

            WaitForReadyState(driver, TimeSpan.FromSeconds(2));
        }

        /// <summary>
        /// Đợi các quảng cáo trên trang web được render hoàn toàn trước khi chụp ảnh
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
                        const iframes = Array.from(document.querySelectorAll('iframe, frame'));
                        for (const fr of iframes){
                          const r = fr.getBoundingClientRect();
                          if (r.width>=minW && r.height>=minH){ score += 2; }
                          try { if (fr.contentWindow) score += 1; } catch(e) {}
                        }
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
                        if (document.readyState==='complete') score += 1;
                        return score;";
                    int score = Convert.ToInt32(js.ExecuteScript(script));
                    if (score == lastScore) stableRounds++; else { stableRounds = 0; lastScore = score; }
                    // Yêu cầu score >= 4 và stable trong ít nhất 3 lần liên tiếp (thay vì 2) để đảm bảo quảng cáo đã load
                    if (score >= 4 && stableRounds >= 3) break;
                }
                catch { }
                Thread.Sleep(400); // Tăng delay giữa các lần kiểm tra để banner có thời gian load
            }
        }

        /// <summary>
        /// Thăm dò xem các phần tử quảng cáo đã xuất hiện trên trang chưa
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

        /// <summary>
        /// Đợi trang web đạt trạng thái ready (đã load xong)
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
    }
}

