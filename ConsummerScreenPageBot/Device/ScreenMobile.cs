using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading;
using ConsummerScreenPageBot.Utils;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Jpeg;

namespace ConsummerScreenPageBot.Device
{
    /// <summary>
    /// Class xử lý chụp màn hình cho Mobile device
    /// Viewport: 375x667 (iPhone), mobile: true
    /// User Agent: Mobile
    /// </summary>
    public static class ScreenMobile
    {
        private static string startupPath = AppDomain.CurrentDomain.BaseDirectory.Replace(@"\bin\Debug\net8.0", @"\");
        private static string LogPath = ConfigurationManager.AppSettings["PATHLOG"] ?? "logs";
        
        // Mobile viewport mặc định: iPhone 12/13 (375x667)
        private const int MOBILE_WIDTH = 375;
        private const int MOBILE_HEIGHT = 667;

        /// <summary>
        /// Xử lý website để chụp ảnh quảng cáo và nội dung trên Mobile
        /// 
        /// Cách thức hoạt động:
        /// 1. Lấy hostname từ URL để tạo thư mục lưu ảnh tương ứng
        /// 2. Xóa toàn bộ ảnh cũ trong thư mục screenshots/<hostname>/mobile để làm sạch
        /// 3. Cuộn trang xuống cuối để kích hoạt lazy-loading (quảng cáo, hình ảnh load chậm)
        ///    - Đợi cho đến khi chiều cao trang không tăng thêm (đã load hết nội dung)
        /// 4. Thăm dò các phần tử có thể là quảng cáo (banner, ads container)
        /// 5. Chụp ảnh theo segment: chia trang thành N phần bằng nhau và chụp từng phần
        ///    - Mỗi ảnh segment sẽ được nén với chất lượng jpegQuality
        ///    - Tự động gửi mỗi ảnh lên queue analyze để AI xử lý
        /// 
        /// Tham số:
        /// - driver: WebDriver đã điều hướng tới trang web với mobile viewport
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
        }

        /// <summary>
        /// Chụp ảnh trang web bằng cách chia thành nhiều segment (phần) và chụp từng phần bằng cách scroll đến từng vị trí - Mobile
        /// 
        /// Viewport Mobile: 375x667, mobile: true
        /// 
        /// Cách thức hoạt động:
        /// 1. Scroll xuống cuối trang để kích hoạt lazy-load và đợi quảng cáo load hoàn toàn
        /// 2. Scroll lên đầu trang để reset vị trí
        /// 3. Tính chiều cao mỗi segment: totalHeight / segmentCount
        /// 4. Với mỗi segment i:
        ///    - Scroll đến vị trí Y = i * sliceHeight (đặt ở giữa viewport)
        ///    - Đợi một chút để đảm bảo content đã render
        ///    - Chụp viewport tại vị trí đó
        ///    - Lưu và nén ảnh với chất lượng jpegQuality
        ///    - Gửi ảnh lên queue analyze để AI xử lý
        /// 5. Scroll lại lên đầu trang sau khi xong
        /// 
        /// Ưu điểm: Mỗi segment là ảnh chụp viewport riêng biệt khi scroll, đảm bảo không bỏ sót nội dung
        /// </summary>
        private static void CaptureSegmentScreenshots(IWebDriver driver, string hostLabel, int segmentCount, long jpegQuality)
        {
            try
            {
                try { Console.WriteLine($"[Mobile Segment] Start capture host={hostLabel}, segments={segmentCount}, quality={jpegQuality}"); } catch { }
                if (segmentCount <= 0) segmentCount = 3;

                var js = (IJavaScriptExecutor)driver;
                // Kích hoạt lazy-load và đợi quảng cáo hiển thị trước khi chụp full
                ScrollToBottomAndEnsureLazyContent(driver, TimeSpan.FromSeconds(10));
                
                // Cuộn về đầu trang để đảm bảo banner top đã load
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                Thread.Sleep(1500); // Đợi banner top load
                
                // Đợi quảng cáo load với thời gian dài hơn
                WaitForAdsLoaded(driver, TimeSpan.FromSeconds(12));
                
                // Delay thêm để đảm bảo tất cả banner quảng cáo đã render hoàn toàn
                Thread.Sleep(2000);
                
                // Bước 3: Lấy chiều cao tổng của trang
                int totalHeight = 3000;
                int viewportHeight = 667; // Mobile viewport mặc định
                try
                {
                    totalHeight = Convert.ToInt32(js.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight, window.innerHeight || 0);"));
                    viewportHeight = Convert.ToInt32(js.ExecuteScript("return window.innerHeight || document.documentElement.clientHeight || 667;"));
                }
                catch { }
                if (totalHeight <= 0) totalHeight = 3000;
                totalHeight = Math.Min(totalHeight, 30000);
                if (viewportHeight <= 0) viewportHeight = 667;

                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel, "mobile");
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                // Bước 4: Tính chiều cao mỗi segment và scroll đến từng vị trí để chụp
                int sliceHeight = (int)Math.Ceiling((double)totalHeight / segmentCount);
                Console.WriteLine($"[Mobile Segment] Total height: {totalHeight}px, Viewport height: {viewportHeight}px, Segment height: {sliceHeight}px");

                // Scroll lên đầu trang trước khi bắt đầu
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                Thread.Sleep(500);

                for (int i = 0; i < segmentCount; i++)
                {
                    // Tính vị trí Y để scroll đến (đặt vị trí segment ở giữa viewport)
                    int targetY = i * sliceHeight;
                    // Điều chỉnh để segment nằm ở giữa viewport
                    int scrollY = Math.Max(0, targetY - (viewportHeight / 2));
                    
                    Console.WriteLine($"[Mobile Segment] Segment {i + 1}/{segmentCount}: Scroll to Y={scrollY} (target segment at Y={targetY})");
                    
                    // Scroll đến vị trí
                    try
                    {
                        js.ExecuteScript($"window.scrollTo(0, {scrollY});");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Mobile Segment] Error scrolling to {scrollY}: {ex.Message}");
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
                        Console.WriteLine($"[Mobile Segment] Error capturing segment {i + 1}: {ex.Message}");
                        continue;
                    }
                    
                    if (segmentBytes == null || segmentBytes.Length == 0)
                    {
                        Console.WriteLine($"[Mobile Segment] WARNING: Segment {i + 1} screenshot is empty");
                        continue;
                    }
                    
                    // Load ảnh và lưu
                    try
                    {
                        using (var ms = new MemoryStream(segmentBytes))
                        using (var img = Image.Load<Rgba32>(ms))
                        {
                            var fileName = $"mobile_split{i+1}_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.jpg";
                            var savePath = Path.Combine(shotsDir, fileName);
                            
                            // Nén ảnh
                            AdCapture.SaveImageCompressed(img, savePath, 1.0, jpegQuality);
                            
                            var compressedBytes = File.ReadAllBytes(savePath);
                            Console.WriteLine($"[Mobile Segment] Saved segment {i + 1}/{segmentCount}: {Path.GetFileName(savePath)} ({compressedBytes.Length} bytes, {img.Width}x{img.Height}px)");
                            
                            // Gửi lên queue
                            Program.TryPublishAnalyze(compressedBytes, img.Width, img.Height);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Mobile Segment] Error processing segment {i + 1}: {ex.Message}");
                        ErrorWriter.WriteLog(LogPath, "MobileSegmentScreenshots.Process", $"Segment {i + 1} - {hostLabel} => {ex}");
                    }
                }
                
                // Scroll lại lên đầu trang sau khi xong
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                try { Console.WriteLine($"[Mobile Segment] Done host={hostLabel}"); } catch { }
                
                // Sau khi chụp segment xong, xử lý tất cả iframe và push vào queue
                Console.WriteLine("[Mobile] Bắt đầu xử lý iframe sau khi chụp segment...");
                Program.ProcessIframesAndPushToQueue(driver, hostLabel);
            }
            catch (Exception ex)
            {
                try { Console.WriteLine($"[Mobile Segment] ERROR {hostLabel}: {ex.Message}"); } catch { }
                ErrorWriter.WriteLog(LogPath, "MobileSegmentScreenshots", ex.ToString());
                TelegramService.PushLogToTelegram($"Mobile SegmentScreenshots Error - {hostLabel}", ex);
            }
        }

        /// <summary>
        /// Xóa tất cả file ảnh cũ trong thư mục screenshots của host cụ thể - Mobile
        /// </summary>
        private static void ClearHostScreenshots(string hostLabel)
        {
            try
            {
                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel, "mobile");
                if (!Directory.Exists(shotsDir)) return;

                var files = Directory.GetFiles(shotsDir);
                foreach (var f in files)
                {
                    try { File.Delete(f); }
                    catch (Exception ex)
                    {
                        ErrorWriter.WriteLog(LogPath, "ClearScreenshots.Delete", $"Mobile {hostLabel} => {f} => {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(LogPath, "ClearScreenshots", $"Mobile {hostLabel} => {ex}");
                TelegramService.PushLogToTelegram($"Mobile ClearScreenshots Error - {hostLabel}", ex);
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
                        const minW=80, minH=30;
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

        /// <summary>
        /// Chụp viewport đầu trang (màn hình cơ sở 1) - dùng cho banner dạng U ngược
        /// </summary>
        private static byte[] CaptureViewportScreenshot(IWebDriver driver, long jpegQuality = 80)
        {
            try
            {
                // Scroll lên đầu trang
                var js = (IJavaScriptExecutor)driver;
                try { js.ExecuteScript("window.scrollTo(0, 0);"); } catch { }
                Thread.Sleep(300);
                
                // Chụp viewport hiện tại
                var chrome = driver as ChromeDriver;
                if (chrome != null)
                {
                    try
                    {
                        var args = new Dictionary<string, object>
                        {
                            { "format", "jpeg" },
                            { "quality", (int)Math.Clamp(jpegQuality, 1, 100) },
                            { "captureBeyondViewport", false }  // Chỉ chụp viewport
                        };
                        var result = chrome.ExecuteCdpCommand("Page.captureScreenshot", args) as IDictionary<string, object>;
                        if (result != null && result.TryGetValue("data", out var dataObj) && dataObj is string base64)
                        {
                            return Convert.FromBase64String(base64);
                        }
                    }
                    catch { }
                }
                
                // Fallback: ITakesScreenshot
                try
                {
                    var shot = ((ITakesScreenshot)driver).GetScreenshot();
                    return shot.AsByteArray;
                }
                catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Mobile] Lỗi khi chụp viewport screenshot: {ex.Message}");
            }
            
            return Array.Empty<byte>();
        }

        /// <summary>
        /// Detect banner dạng chữ U ngược: banner trên + banner trái + banner phải (Mobile)
        /// Mobile viewport nhỏ nên điều chỉnh ngưỡng cho phù hợp
        /// </summary>
        private static bool DetectUShapedBannerLayout(IWebDriver driver, List<(int Left, int Top, int Right, int Bottom)> bannerRects, int pageWidth, out (int Left, int Top, int Right, int Bottom)? uShapeBounds)
        {
            uShapeBounds = null;
            try
            {
                // Phát hiện banner ở 3 vị trí: top, left, right
                var topBanners = new List<(int Left, int Top, int Right, int Bottom)>();
                var leftBanners = new List<(int Left, int Top, int Right, int Bottom)>();
                var rightBanners = new List<(int Left, int Top, int Right, int Bottom)>();
                
                foreach (var rect in bannerRects)
                {
                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;
                    
                    // Top banner: ở trên cùng (top < 400px), rộng >= 60% page width (mobile viewport nhỏ)
                    if (rect.Top < 400 && width >= pageWidth * 0.6)
                    {
                        topBanners.Add(rect);
                    }
                    // Left sidebar: ở bên trái (left < 150px), không phải top banner (mobile viewport nhỏ)
                    else if (rect.Left < 150 && rect.Top >= 150)
                    {
                        leftBanners.Add(rect);
                    }
                    // Right sidebar: ở bên phải (right > pageWidth - 150px), không phải top banner
                    else if (rect.Right > pageWidth - 150 && rect.Top >= 150)
                    {
                        rightBanners.Add(rect);
                    }
                }
                
                // Kiểm tra có đủ 3 loại banner để tạo hình U ngược
                if (topBanners.Count > 0 && (leftBanners.Count > 0 || rightBanners.Count > 0))
                {
                    // Lấy top banner đầu tiên
                    var topBanner = topBanners.OrderBy(b => b.Top).First();
                    
                    // Tính bounds của U shape: từ top banner đến bottom của left/right banner
                    int uTop = topBanner.Top;
                    int uLeft = leftBanners.Count > 0 ? 0 : topBanner.Left;
                    int uRight = pageWidth;
                    int leftBottom = leftBanners.Count > 0 ? leftBanners.Max(b => b.Bottom) : 0;
                    int rightBottom = rightBanners.Count > 0 ? rightBanners.Max(b => b.Bottom) : 0;
                    int uBottom = Math.Max(topBanner.Bottom, Math.Max(leftBottom, rightBottom));
                    
                    // Chỉ coi là U shape nếu bottom >= 600px (mobile viewport nhỏ hơn desktop)
                    if (uBottom >= 600)
                    {
                        uShapeBounds = (uLeft, uTop, uRight, uBottom);
                        Console.WriteLine($"[Mobile] Detect U-shape banner: top={uTop}, bottom={uBottom}, left={uLeft}, right={uRight}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Mobile] Lỗi khi detect U-shape banner: {ex.Message}");
            }
            
            return false;
        }

        /// <summary>
        /// Gộp các banner gần nhau thành 1 ảnh để giảm số lượng request (tránh rate limit Gemini) - Mobile
        /// 
        /// Logic:
        /// - Tính khoảng cách giữa các banner (theo cả X và Y)
        /// - Nếu khoảng cách < mergeDistance: gộp thành 1 rect lớn hơn
        /// - Giúp giảm số lượng request gửi vào Gemini API
        /// - Mobile: mergeDistance nhỏ hơn (100px) vì viewport nhỏ
        /// </summary>
        private static List<(int Left, int Top, int Right, int Bottom)> MergeNearbyBanners(
            List<(int Left, int Top, int Right, int Bottom)> bannerRects, 
            int mergeDistance = 100)
        {
            var merged = new List<(int Left, int Top, int Right, int Bottom)>();
            var processed = new HashSet<int>();
            
            try
            {
                // Sắp xếp banner theo vị trí (top trước, sau đó left)
                var sortedBanners = bannerRects
                    .Select((rect, index) => new { Rect = rect, Index = index })
                    .OrderBy(x => x.Rect.Top)
                    .ThenBy(x => x.Rect.Left)
                    .ToList();
                
                for (int i = 0; i < sortedBanners.Count; i++)
                {
                    if (processed.Contains(sortedBanners[i].Index)) continue;
                    
                    var current = sortedBanners[i].Rect;
                    var group = new List<(int Left, int Top, int Right, int Bottom)> { current };
                    processed.Add(sortedBanners[i].Index);
                    
                    // Tìm các banner gần nhau (trong phạm vi mergeDistance)
                    for (int j = i + 1; j < sortedBanners.Count; j++)
                    {
                        if (processed.Contains(sortedBanners[j].Index)) continue;
                        
                        var other = sortedBanners[j].Rect;
                        
                        // Tính khoảng cách ngắn nhất giữa 2 banner (từ cạnh đến cạnh)
                        int distanceX = 0;
                        int distanceY = 0;
                        
                        if (other.Right < current.Left)
                            distanceX = current.Left - other.Right;
                        else if (current.Right < other.Left)
                            distanceX = other.Left - current.Right;
                        // Nếu overlap hoặc gần nhau trên trục X: distanceX = 0
                        
                        if (other.Bottom < current.Top)
                            distanceY = current.Top - other.Bottom;
                        else if (current.Bottom < other.Top)
                            distanceY = other.Top - current.Bottom;
                        // Nếu overlap hoặc gần nhau trên trục Y: distanceY = 0
                        
                        // Khoảng cách tổng (Manhattan distance)
                        int totalDistance = distanceX + distanceY;
                        
                        // Nếu khoảng cách < mergeDistance hoặc overlap: gộp lại
                        if (totalDistance < mergeDistance || distanceX == 0 || distanceY == 0)
                        {
                            group.Add(other);
                            processed.Add(sortedBanners[j].Index);
                        }
                    }
                    
                    // Gộp group thành 1 rect lớn hơn
                    if (group.Count > 1)
                    {
                        int minLeft = group.Min(b => b.Left);
                        int minTop = group.Min(b => b.Top);
                        int maxRight = group.Max(b => b.Right);
                        int maxBottom = group.Max(b => b.Bottom);
                        merged.Add((minLeft, minTop, maxRight, maxBottom));
                        Console.WriteLine($"[Mobile] Gộp {group.Count} banner gần nhau: ({group[0].Left},{group[0].Top}) -> ({minLeft},{minTop}) ~ ({maxRight},{maxBottom})");
                    }
                    else
                    {
                        merged.Add(current);
                    }
                }
                
                return merged;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Mobile] Lỗi khi merge banner: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "MergeNearbyBanners", ex.ToString());
                return bannerRects; // Fallback: trả về danh sách gốc
            }
        }

        /// <summary>
        /// Detect các vùng banner trên trang và trả về danh sách tọa độ (left, top, right, bottom) - Mobile
        /// </summary>
        private static List<(int Left, int Top, int Right, int Bottom)> DetectBannerRegions(IWebDriver driver)
        {
            var bannerRects = new List<(int Left, int Top, int Right, int Bottom)>();
            try
            {
                var js = (IJavaScriptExecutor)driver;
                
                // Lấy scrollY để tính tọa độ tuyệt đối
                long scrollY = 0;
                try
                {
                    scrollY = Convert.ToInt64(js.ExecuteScript("return window.scrollY || window.pageYOffset || document.documentElement.scrollTop || 0;"));
                }
                catch { }
                
                // Detect iframes (thường là quảng cáo) - Mobile có ngưỡng nhỏ hơn
                var iframeScript = @"
                    const iframes = Array.from(document.querySelectorAll('iframe, frame'));
                    const minW = 80, minH = 30;  // Mobile: giảm ngưỡng width xuống 80px
                    const y = window.scrollY || window.pageYOffset || 0;
                    const rects = [];
                    
                    for (const fr of iframes) {
                        try {
                            const r = fr.getBoundingClientRect();
                            const w = Math.round(r.width);
                            const h = Math.round(r.height);
                            
                            if (w >= minW && h >= minH) {
                                const style = window.getComputedStyle(fr);
                                if (style.display !== 'none' && style.visibility !== 'hidden') {
                                    const left = Math.round(r.left);
                                    const top = Math.round(r.top + y);
                                    const right = Math.round(r.right);
                                    const bottom = Math.round(r.bottom + y);
                                    rects.push([left, top, right, bottom]);
                                }
                            }
                        } catch(e) {}
                    }
                    return rects;
                ";
                
                var iframeRects = js.ExecuteScript(iframeScript) as System.Collections.IEnumerable;
                if (iframeRects != null)
                {
                    foreach (var item in iframeRects)
                    {
                        var rect = item as System.Collections.IList;
                        if (rect != null && rect.Count >= 4)
                        {
                            int left = Convert.ToInt32(rect[0]);
                            int top = Convert.ToInt32(rect[1]);
                            int right = Convert.ToInt32(rect[2]);
                            int bottom = Convert.ToInt32(rect[3]);
                            if (right > left && bottom > top)
                            {
                                bannerRects.Add((left, top, right, bottom));
                            }
                        }
                    }
                }
                
                // Detect banner elements bằng CSS selector - Mobile có ngưỡng nhỏ hơn
                var adSelectors = AdCapture.GetCommonAdSelectors();
                var adSelectorScript = @"
                    const selectors = arguments[0];
                    const minW = 80, minH = 30;  // Mobile: giảm ngưỡng width xuống 80px
                    const y = window.scrollY || window.pageYOffset || 0;
                    const rects = [];
                    const processed = new Set();
                    
                    for (const sel of selectors) {
                        try {
                            const elements = Array.from(document.querySelectorAll(sel));
                            for (const el of elements) {
                                const r = el.getBoundingClientRect();
                                const w = Math.round(r.width);
                                const h = Math.round(r.height);
                                
                                if (w >= minW && h >= minH) {
                                    const style = window.getComputedStyle(el);
                                    if (style.display !== 'none' && style.visibility !== 'hidden') {
                                        const left = Math.round(r.left);
                                        const top = Math.round(r.top + y);
                                        const right = Math.round(r.right);
                                        const bottom = Math.round(r.bottom + y);
                                        
                                        // Tránh trùng với iframe đã detect
                                        let isDuplicate = false;
                                        for (const existing of rects) {
                                            if (Math.abs(existing[0] - left) < 10 && 
                                                Math.abs(existing[1] - top) < 10 &&
                                                Math.abs(existing[2] - right) < 10 &&
                                                Math.abs(existing[3] - bottom) < 10) {
                                                isDuplicate = true;
                                                break;
                                            }
                                        }
                                        
                                        if (!isDuplicate) {
                                            rects.push([left, top, right, bottom]);
                                        }
                                    }
                                }
                            }
                        } catch(e) {}
                    }
                    
                    return rects;
                ";
                
                var adRects = js.ExecuteScript(adSelectorScript, adSelectors) as System.Collections.IEnumerable;
                if (adRects != null)
                {
                    foreach (var item in adRects)
                    {
                        var rect = item as System.Collections.IList;
                        if (rect != null && rect.Count >= 4)
                        {
                            int left = Convert.ToInt32(rect[0]);
                            int top = Convert.ToInt32(rect[1]);
                            int right = Convert.ToInt32(rect[2]);
                            int bottom = Convert.ToInt32(rect[3]);
                            if (right > left && bottom > top)
                            {
                                bannerRects.Add((left, top, right, bottom));
                            }
                        }
                    }
                }
                
                // Sắp xếp banner theo vị trí top (từ trên xuống)
                bannerRects = bannerRects.OrderBy(r => r.Top).ToList();
                
                Console.WriteLine($"[Mobile] Detect được {bannerRects.Count} vùng banner");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Mobile] Lỗi khi detect banner: {ex.Message}");
                ErrorWriter.WriteLog(LogPath, "DetectBannerRegions", ex.ToString());
            }
            
            return bannerRects;
        }
    }
}

