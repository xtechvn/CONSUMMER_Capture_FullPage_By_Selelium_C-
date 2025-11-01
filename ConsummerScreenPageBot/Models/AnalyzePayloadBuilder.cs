using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ConsummerScreenPageBot.Models
{
	public static class AnalyzePayloadBuilder
	{
		public static byte[] BuildAnalyzeBody(byte[] imageBytes, JObject? jobParamsSnapshot, string analyzePublishRaw)
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

			merged["screenshot_base64"] = System.Convert.ToBase64String(imageBytes);
			var json = merged.ToString(Formatting.None);
			return Encoding.UTF8.GetBytes(json);
		}
	}
}


