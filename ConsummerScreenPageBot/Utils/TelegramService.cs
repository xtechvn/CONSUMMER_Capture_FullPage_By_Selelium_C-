using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ConsummerScreenPageBot.Utils
{
    public class TelegramService
    {
        private const string TELEGRAM_BOT_TOKEN = "7633683325:AAEPqbTQRaifoz_dVectxS180j5fdXHHip0";
        private const string TELEGRAM_CHAT_ID = "807032654";
        private const string TELEGRAM_API_URL = "https://api.telegram.org/bot{0}/sendMessage";
        private const string TELEGRAM_LOG_PATH = "logs";
        
        private static readonly HttpClient httpClient = new HttpClient();

        /// <summary>
        /// Push log message to Telegram
        /// </summary>
        /// <param name="messageLog">Message log to send</param>
        /// <returns>True if successful, False otherwise</returns>
        public static async Task<bool> PushLogToTelegramAsync(string messageLog)
        {
            try
            {
                var url = string.Format(TELEGRAM_API_URL, TELEGRAM_BOT_TOKEN);
                
                var payload = new
                {
                    chat_id = TELEGRAM_CHAT_ID,
                    text = messageLog,
                    parse_mode = "HTML"
                };

                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync(url, content);
                
                if (response.IsSuccessStatusCode)
                {
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<TelegramResponse>(responseContent);
                    
                    return result?.ok == true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                // Log error to console and file
                try
                {
                    Console.WriteLine($"Error sending message to Telegram: {ex.Message}");
                    ConsummerScreenPageBot.ErrorWriter.WriteLog(TELEGRAM_LOG_PATH, "TelegramService.PushLogToTelegramAsync", 
                        $"Failed to send message. Error: {ex}");
                }
                catch
                {
                    // Ultimate fallback - only console
                    Console.WriteLine($"Critical: Failed to log Telegram error: {ex.Message}");
                }
                return false;
            }
        }

        /// <summary>
        /// Push log message to Telegram (synchronous version)
        /// </summary>
        /// <param name="messageLog">Message log to send</param>
        /// <returns>True if successful, False otherwise</returns>
        public static bool PushLogToTelegram(string messageLog)
        {
            try
            {
                var url = string.Format(TELEGRAM_API_URL, TELEGRAM_BOT_TOKEN);
                
                var payload = new
                {
                    chat_id = TELEGRAM_CHAT_ID,
                    text = messageLog,
                    parse_mode = "HTML"
                };

                var json = JsonConvert.SerializeObject(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = httpClient.PostAsync(url, content).Result;
                
                if (response.IsSuccessStatusCode)
                {
                    var responseContent = response.Content.ReadAsStringAsync().Result;
                    var result = JsonConvert.DeserializeObject<TelegramResponse>(responseContent);
                    
                    return result?.ok == true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                // Log error to console and file
                try
                {
                    Console.WriteLine($"Error sending message to Telegram: {ex.Message}");
                    ConsummerScreenPageBot.ErrorWriter.WriteLog(TELEGRAM_LOG_PATH, "TelegramService.PushLogToTelegram", 
                        $"Failed to send message. Error: {ex}");
                }
                catch
                {
                    // Ultimate fallback - only console
                    Console.WriteLine($"Critical: Failed to log Telegram error: {ex.Message}");
                }
                return false;
            }
        }

        /// <summary>
        /// Push log message to Telegram with exception details
        /// </summary>
        /// <param name="messageLog">Message log to send</param>
        /// <param name="exception">Exception object</param>
        /// <returns>True if successful, False otherwise</returns>
        public static async Task<bool> PushLogToTelegramAsync(string messageLog, Exception exception)
        {
            try
            {
                var fullMessage = $"{messageLog}\n\n<b>Exception:</b>\n<pre>{exception}</pre>";
                return await PushLogToTelegramAsync(fullMessage);
            }
            catch (Exception ex)
            {
                try
                {
                    Console.WriteLine($"Error in PushLogToTelegramAsync with exception: {ex.Message}");
                    ConsummerScreenPageBot.ErrorWriter.WriteLog(TELEGRAM_LOG_PATH, "TelegramService.PushLogToTelegramAsync.Ex", 
                        $"Failed to send message with exception. Error: {ex}");
                }
                catch
                {
                    Console.WriteLine($"Critical: Failed to log Telegram error: {ex.Message}");
                }
                return false;
            }
        }

        /// <summary>
        /// Push log message to Telegram with exception details (synchronous version)
        /// </summary>
        /// <param name="messageLog">Message log to send</param>
        /// <param name="exception">Exception object</param>
        /// <returns>True if successful, False otherwise</returns>
        public static bool PushLogToTelegram(string messageLog, Exception exception)
        {
            try
            {
                var fullMessage = $"{messageLog}\n\n<b>Exception:</b>\n<pre>{exception}</pre>";
                return PushLogToTelegram(fullMessage);
            }
            catch (Exception ex)
            {
                try
                {
                    Console.WriteLine($"Error in PushLogToTelegram with exception: {ex.Message}");
                    ConsummerScreenPageBot.ErrorWriter.WriteLog(TELEGRAM_LOG_PATH, "TelegramService.PushLogToTelegram.Ex", 
                        $"Failed to send message with exception. Error: {ex}");
                }
                catch
                {
                    Console.WriteLine($"Critical: Failed to log Telegram error: {ex.Message}");
                }
                return false;
            }
        }

        private class TelegramResponse
        {
            public bool ok { get; set; }
            public string? description { get; set; }
        }
    }
}

