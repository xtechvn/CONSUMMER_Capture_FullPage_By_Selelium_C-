using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Jpeg;

namespace ConsummerScreenPageBot
{
    public static class AdCapture
    {
		// Lưu ảnh nén (cross-platform) với ImageSharp: mặc định giữ nguyên kích thước, chỉ nén JPEG theo chất lượng
		public static void SaveImageCompressed(Image<Rgba32> source, string savePath, double scaleFactor = 1.0, long jpegQuality = 80)
		{
			try
			{
				int targetW = Math.Max(1, (int)Math.Round(source.Width * scaleFactor));
				int targetH = Math.Max(1, (int)Math.Round(source.Height * scaleFactor));

				using (var clone = source.Clone(ctx => ctx.Resize(new ResizeOptions
				{
					Mode = ResizeMode.Stretch,
					Size = new Size(targetW, targetH)
				})))
				{
					var encoder = new JpegEncoder { Quality = (int)Math.Clamp(jpegQuality, 1, 100) };
					clone.Save(savePath, encoder);
				}
			}
			catch (Exception ex)
			{
				ErrorWriter.WriteLog("logs", "SaveImageCompressed", ex.ToString());
			}
		}

		// Tiện ích: nén và lưu từ mảng byte (ảnh gốc)
		// scaleFactor dùng để điều chỉnh tỉ lệ co của ảnh gốc khi lưu: ảnh đầu ra sẽ được resize về (scaleFactor * chiều rộng, scaleFactor * chiều cao).
		// Việc chọn 0.5 nghĩa là thu nhỏ ảnh còn 50% kích thước gốc về cả chiều rộng và chiều cao. 
		// Điều này giúp giảm đáng kể dung lượng file và tăng hiệu suất tải, xử lý khi lưu hoặc phân phối ảnh, 
		// đồng thời vẫn giữ được hình ảnh rõ ràng đủ dùng cho việc nhận diện quảng cáo.
		// Ngoài ra, giảm kích thước ảnh giúp tiết kiệm băng thông mạng, dung lượng lưu trữ 
		// và tăng tốc độ gửi ảnh qua các hệ thống khác (như RabbitMQ, API, ...).
		/// <summary>
		/// Tham số jpegQuality xác định mức chất lượng của ảnh đầu ra khi lưu dưới định dạng JPEG.
		/// Giá trị này là số nguyên trong khoảng từ 1 đến 100 (mặc định sử dụng 80), 
		/// với số càng lớn thì giữ được nhiều chi tiết cũng như chất lượng hình ảnh cao hơn,
		/// nhưng dung lượng file cũng sẽ lớn hơn. Nếu giảm jpegQuality xuống mức thấp hơn,
		/// ảnh sẽ bị nén mạnh hơn, giảm kích thước file nhưng có thể bị mất chi tiết hoặc xuất hiện hiện tượng nén.
		/// Thông thường, chất lượng từ 70 đến 90 là phù hợp cho mục đích nhận diện quảng cáo mà vẫn tiết kiệm dung lượng.
		/// </summary>
		public static void SaveJpegCompressedFromBytes(byte[] sourceBytes, string savePath, double scaleFactor = 1.0, long jpegQuality = 80)
		{
			try
			{
				using (var ms = new MemoryStream(sourceBytes))
				using (var img = Image.Load<Rgba32>(ms))
				{
					SaveImageCompressed(img, savePath, scaleFactor, jpegQuality);
				}
			}
			catch (Exception ex)
			{
				ErrorWriter.WriteLog("logs", "SaveJpegCompressedFromBytes", ex.ToString());
			}
		}

        // Bộ selector chung áp dụng cho nhiều trang tin
        public static string[] GetCommonAdSelectors()
        {
            return new[]
            {
                "[id*='ad'], [id*='ads'], [id*='advert'], [id*='banner']",
                "[class*='ad '], [class*=' ad'], [class*='ad-'], [class*='-ad'], [class*='ads'], [class*='advert'], [class*='banner']",
                ".gpt-ad, .gpt-unit, .gpt-slot, .dfp-ad, .dfp-slot, .ad-slot, .ad-container, .ad-wrapper, .ad__container, .ad__slot, .adsbygoogle, .google-auto-placed",
                ".sponsor, .sponsored, .promoted, .native-ad, .in-content-ad, .article-ad, .article-advertisement",
                ".header-ad, .footer-ad, .sidebar-ad, .sticky-ad, .floating-ad, #footer-ad, .footer-banner, .footer-advertisement",
                "[data-ad], [data-ad-slot], [data-ad-unit], [data-google-query-id], [data-ez-name]"
            };
        }

        // Quét và chụp theo selector ở tài liệu hiện tại
        public static void CaptureBySelectors(IWebDriver driver, string[] selectors, string hostLabel, string startupPath, string logPath, long jpegQuality = 80)
        {
            int saved = 0;
            foreach (var css in selectors)
            {
                try
                {
                    var elements = driver.FindElements(By.CssSelector(css));
                    int localIndex = 0;
                    foreach (var el in elements.Take(10))
                    {
                        try
                        {
                            localIndex++;
                            try { ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", el); } catch { }
                            System.Threading.Thread.Sleep(120);

                            int measuredW = 0, measuredH = 0;
                            try
                            {
                                var rect = ((IJavaScriptExecutor)driver).ExecuteScript(
                                    "var r=arguments[0].getBoundingClientRect(); return [Math.round(r.width), Math.round(r.height)];",
                                    el) as System.Collections.IList;
                                if (rect != null && rect.Count >= 2)
                                {
                                    measuredW = Convert.ToInt32(rect[0]);
                                    measuredH = Convert.ToInt32(rect[1]);
                                }
                            }
                            catch { }

                            if (measuredW <= 0 || measuredH <= 0)
                            {
                                measuredW = el.Size.Width;
                                measuredH = el.Size.Height;
                            }

                            Console.WriteLine($"Candidate [{css}] => {measuredW}x{measuredH}");
                            if (!el.Displayed || measuredH < 30 || measuredW < 120) continue;

                            var label = BuildElementLabel(el, css);
                            SaveElementScreenshot(el, hostLabel, label, measuredW, measuredH, localIndex, startupPath, logPath, jpegQuality);
                            saved++;
                        }
                        catch (Exception inner)
                        {
                            ErrorWriter.WriteLog(logPath, "CaptureElement", inner.ToString());
                        }
                    }
                }
                catch (NoSuchElementException) { }
                catch (Exception ex)
                {
                    ErrorWriter.WriteLog(logPath, "QuerySelector", ex.ToString());
                }
            }
            Console.WriteLine($"Saved {saved} banner screenshots for {hostLabel}");
        }

        // Chụp trực tiếp các iframe/frame như một phần tử riêng lẻ
        public static void CaptureAdIframes(IWebDriver driver, string hostLabel, string startupPath, string logPath, long jpegQuality = 80)
        {
            try
            {
                var frames = driver.FindElements(By.CssSelector("iframe, frame")).Take(20).ToList();
                int idx = 0;
                foreach (var fr in frames)
                {
                    idx++;
                    try
                    {
                        try { ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", fr); } catch { }
                        System.Threading.Thread.Sleep(120);

                        int measuredW = 0, measuredH = 0;
                        try
                        {
                            var rect = ((IJavaScriptExecutor)driver).ExecuteScript(
                                "var r=arguments[0].getBoundingClientRect(); return [Math.round(r.width), Math.round(r.height)];",
                                fr) as System.Collections.IList;
                            if (rect != null && rect.Count >= 2)
                            {
                                measuredW = Convert.ToInt32(rect[0]);
                                measuredH = Convert.ToInt32(rect[1]);
                            }
                        }
                        catch { }

                        if (measuredW <= 0 || measuredH <= 0)
                        {
                            measuredW = fr.Size.Width;
                            measuredH = fr.Size.Height;
                        }

                        if (measuredW < 120 || measuredH < 30) continue;

                        var label = BuildElementLabel(fr, "iframe");
                        SaveElementScreenshot(fr, hostLabel, label, measuredW, measuredH, idx, startupPath, logPath, jpegQuality);
                    }
                    catch (Exception ex)
                    {
                        ErrorWriter.WriteLog(logPath, "CaptureAdIframes", ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(logPath, "CaptureAdIframes.Root", ex.ToString());
            }
        }

		private static void SaveElementScreenshot(IWebElement element, string hostLabel, string elementLabel, int measuredWidth, int measuredHeight, int index, string startupPath, string logPath, long jpegQuality)
        {
            try
            {
                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                var driver = ((IWrapsDriver)element).WrappedDriver;
                var fullShot = ((ITakesScreenshot)driver).GetScreenshot();

				using (var ms = new MemoryStream(fullShot.AsByteArray))
				using (var img = Image.Load<Rgba32>(ms))
				{
					var rect = new Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);
					int rx = Math.Max(0, rect.X);
					int ry = Math.Max(0, rect.Y);
					int rw = Math.Min(rect.Width, img.Width - rx);
					int rh = Math.Min(rect.Height, img.Height - ry);
					if (rw <= 0 || rh <= 0)
					{
						var fallbackName = $"{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.jpg";
						var fallbackPath = Path.Combine(shotsDir, fallbackName);
						SaveImageCompressed(img, fallbackPath, 1.0, jpegQuality);
						return;
					}

					using (var crop = img.Clone(ctx => ctx.Crop(new Rectangle(rx, ry, rw, rh))))
					{
						var sizePart = $"{measuredWidth}x{measuredHeight}";
						var safeLabel = string.IsNullOrWhiteSpace(elementLabel) ? "element" : elementLabel;
						var fileName = $"{safeLabel}_{sizePart}_{index}.jpg";
						var savePath = Path.Combine(shotsDir, fileName);
						SaveImageCompressed(crop, savePath, 1.0, jpegQuality);
						
						try
						{
							var bytes = File.ReadAllBytes(savePath);
							try { Console.WriteLine($"[AdCapture] Publish analyze: {hostLabel} -> {Path.GetFileName(savePath)} ({bytes.Length} bytes)"); } catch { }
							Program.TryPublishAnalyze(bytes);
						}
						catch { }
					}
				}
            }
            catch (Exception ex)
            {
                ErrorWriter.WriteLog(logPath, "SaveScreenshot", ex.ToString());
            }
        }

        private static string BuildElementLabel(IWebElement element, string css)
        {
            try
            {
                var id = element.GetAttribute("id");
                if (!string.IsNullOrWhiteSpace(id)) return SanitizeForFileName(id);

                var cls = element.GetAttribute("class");
                if (!string.IsNullOrWhiteSpace(cls))
                {
                    var first = cls.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(first)) return SanitizeForFileName(first);
                }

                return SanitizeForFileName(css.Replace(" ", "_").Replace(",", "_"));
            }
            catch { return "element"; }
        }

        private static string SanitizeForFileName(string input)
        {
            var sb = new StringBuilder();
            foreach (var ch in input)
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9') || ch == '-' || ch == '_')
                {
                    sb.Append(ch);
                }
                else if (ch == ' ')
                {
                    sb.Append('_');
                }
            }
            var value = sb.ToString();
            if (string.IsNullOrWhiteSpace(value)) value = "element";
            return value.Length > 60 ? value.Substring(0, 60) : value;
        }
    }
}


