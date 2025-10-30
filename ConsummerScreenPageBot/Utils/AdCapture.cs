using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace ConsummerScreenPageBot
{
    public static class AdCapture
    {
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
        public static void CaptureBySelectors(IWebDriver driver, string[] selectors, string hostLabel, string startupPath, string logPath)
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
                            SaveElementScreenshot(el, hostLabel, label, measuredW, measuredH, localIndex, startupPath, logPath);
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
        public static void CaptureAdIframes(IWebDriver driver, string hostLabel, string startupPath, string logPath)
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
                        SaveElementScreenshot(fr, hostLabel, label, measuredW, measuredH, idx, startupPath, logPath);
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

        private static void SaveElementScreenshot(IWebElement element, string hostLabel, string elementLabel, int measuredWidth, int measuredHeight, int index, string startupPath, string logPath)
        {
            try
            {
                var shotsDir = Path.Combine(startupPath, "screenshots", hostLabel);
                if (!Directory.Exists(shotsDir)) Directory.CreateDirectory(shotsDir);

                var driver = ((IWrapsDriver)element).WrappedDriver;
                var fullShot = ((ITakesScreenshot)driver).GetScreenshot();

                using (var ms = new MemoryStream(fullShot.AsByteArray))
                using (var bmp = new Bitmap(ms))
                {
                    var rect = new Rectangle(element.Location, element.Size);
                    rect.X = Math.Max(0, rect.X);
                    rect.Y = Math.Max(0, rect.Y);
                    rect.Width = Math.Min(rect.Width, bmp.Width - rect.X);
                    rect.Height = Math.Min(rect.Height, bmp.Height - rect.Y);
                    if (rect.Width <= 0 || rect.Height <= 0)
                    {
                        var fallbackName = $"{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N").Substring(0,6)}.png";
                        var fallbackPath = Path.Combine(shotsDir, fallbackName);
                        fullShot.SaveAsFile(fallbackPath);
                        return;
                    }

                    using (var crop = bmp.Clone(rect, bmp.PixelFormat))
                    {
                        var sizePart = $"{measuredWidth}x{measuredHeight}";
                        var safeLabel = string.IsNullOrWhiteSpace(elementLabel) ? "element" : elementLabel;
                        var fileName = $"{safeLabel}_{sizePart}_{index}.png";
                        var savePath = Path.Combine(shotsDir, fileName);
                        crop.Save(savePath, ImageFormat.Png);
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


