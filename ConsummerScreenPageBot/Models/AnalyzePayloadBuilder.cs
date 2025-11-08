using System.Text;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ConsummerScreenPageBot.Models
{
	public static class AnalyzePayloadBuilder
	{
		public static byte[] BuildAnalyzeBody(byte[] imageBytes, JObject? jobParamsSnapshot, string analyzePublishRaw, int? width = null, int? height = null)
		{
			if (analyzePublishRaw == "1")
			{
				return imageBytes;
			}

		JObject merged;
		if (jobParamsSnapshot != null)
		{
			try { merged = (JObject)jobParamsSnapshot.DeepClone(); }
			catch { merged = new JObject(); }
		}
		else
		{
			merged = new JObject();
		}

		
		
		// Debug: Log input parameters
		System.Console.WriteLine($"[AnalyzePayload] Input - width.HasValue: {width.HasValue}, width: {(width.HasValue ? width.Value.ToString() : "null")}, height.HasValue: {height.HasValue}, height: {(height.HasValue ? height.Value.ToString() : "null")}");
		
		// Xóa width/height cũ nếu có trong jobParamsSnapshot để tránh conflict
		if (merged.ContainsKey("width"))
		{
			merged.Remove("width");
			System.Console.WriteLine($"[AnalyzePayload] Removed old 'width' from snapshot");
		}
		if (merged.ContainsKey("height"))
		{
			merged.Remove("height");
			System.Console.WriteLine($"[AnalyzePayload] Removed old 'height' from snapshot");
		}
		
		// Thêm width và height nếu có (luôn thêm nếu có giá trị, kể cả 0)
		// Đảm bảo width/height luôn được thêm vào sau screenshot_base64 để không bị ghi đè
		if (width.HasValue)
		{
			merged["width"] = width.Value;
			System.Console.WriteLine($"[AnalyzePayload] Added width: {width.Value}");
		}
		else
		{
			System.Console.WriteLine($"[AnalyzePayload] Warning: width is null or not provided");
		}
		
		if (height.HasValue)
		{
			merged["height"] = height.Value;
			System.Console.WriteLine($"[AnalyzePayload] Added height: {height.Value}");
		}
		else
		{
			System.Console.WriteLine($"[AnalyzePayload] Warning: height is null or not provided");
		}
		merged["screenshot_base64"] = System.Convert.ToBase64String(imageBytes);
		// Debug: Log payload để kiểm tra
		try
		{
			var hasWidth = merged.ContainsKey("width");
			var hasHeight = merged.ContainsKey("height");
			var widthValue = hasWidth ? merged["width"]?.ToString() : "null";
			var heightValue = hasHeight ? merged["height"]?.ToString() : "null";
			System.Console.WriteLine($"[AnalyzePayload] Final check - Width in JSON: {hasWidth} (value: {widthValue}), Height in JSON: {hasHeight} (value: {heightValue})");
			
			// Log một phần JSON để verify - kiểm tra cả đầu và cuối
			var sampleJson = merged.ToString(Formatting.None);
			if (sampleJson.Length > 400)
			{
				System.Console.WriteLine($"[AnalyzePayload] JSON sample (first 200 chars): {sampleJson.Substring(0, 200)}...");
				// Log phần cuối để xem width/height có ở đó không
				var lastPart = sampleJson.Substring(Math.Max(0, sampleJson.Length - 200));
				System.Console.WriteLine($"[AnalyzePayload] JSON sample (last 200 chars): ...{lastPart}");
			}
			else
			{
				System.Console.WriteLine($"[AnalyzePayload] Full JSON: {sampleJson}");
			}
			
			// Kiểm tra xem width/height có thực sự trong JSON string không
			var jsonString = merged.ToString(Formatting.None);
			var hasWidthInString = jsonString.Contains("\"width\"");
			var hasHeightInString = jsonString.Contains("\"height\"");
			System.Console.WriteLine($"[AnalyzePayload] String check - Contains 'width': {hasWidthInString}, Contains 'height': {hasHeightInString}");
			
			// Log tất cả keys trong merged để debug
			var allKeys = string.Join(", ", merged.Properties().Select(p => p.Name));
			System.Console.WriteLine($"[AnalyzePayload] All keys in JSON: {allKeys}");
		}
		catch (Exception ex)
		{
			System.Console.WriteLine($"[AnalyzePayload] Error in debug log: {ex.Message}");
		}
			
			var json = merged.ToString(Formatting.None);
			return Encoding.UTF8.GetBytes(json);
		}
	}
}


