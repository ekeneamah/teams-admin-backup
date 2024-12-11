using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;

namespace TeamsChatBackup
{
    class Program
    {
        private static IConfigurationRoot configuration;

        // Configuration variables
        private static string clientId;
        private static string tenantId;
        private static string clientSecret;
        private static string graphEndpoint = "https://graph.microsoft.com/v1.0";
        private static HttpClient httpClient = new HttpClient();
        private static string token;
        private static DateTime tokenExpiresOn;
        private static int userChatCount = 0;


        static async Task Main(string[] args)
        {
            Console.WriteLine("Starting Teams Chat Backup...");

            var builder = new ConfigurationBuilder()
          .SetBasePath(AppContext.BaseDirectory) // Set the base path to where the app runs
          .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true); // Add appsettings.json
            configuration = builder.Build();

            // Read configuration values
             clientId = configuration["AzureConfig:ClientId"];
            tenantId = configuration["AzureConfig:TenantId"];
            clientSecret = configuration["AzureConfig:ClientSecret"];
            // Ensure args are provided
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: TeamsChatBackup <BackupPath> [Days]");
                //return;
            }

            string backupPath = configuration["BackupConfig:BackupPath"];
            int days =  int.Parse(configuration["BackupConfig:Days"]);

            // Authenticate and get token
            // token = await GetGraphTokenAsync();
            token = await GetApplicationTokenAsync(clientId, clientSecret, tenantId);

            // Fetch users
            var users = await FetchUsersAsync();

            // Create backup directory
            string timestamp = DateTime.UtcNow.ToString("yyyy_MM_dd_HHmm");
            string rootDirectory = Path.Combine(backupPath, $"TeamsChatBackup_{timestamp}");
            Directory.CreateDirectory(rootDirectory);

            foreach (var user in users)
            {
               
                    await BackupUserChatsAsync(user, rootDirectory, days);
                
            }

            Console.WriteLine("Backup completed!");
        }

        private static async Task<string> GetGraphTokenAsync()
        {
            var app = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token"))
                .Build();

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

            return result.AccessToken;
        }

        public static async Task<string> GetApplicationTokenAsync(string clientId, string clientSecret, string tenantId)
        {
            // If the token is still valid, return it
            if (token != null && DateTime.UtcNow < tokenExpiresOn)
            {
                return token;
            }

            // Construct the token request URL
            string uri = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            // Construct the request body
            var requestBody = new FormUrlEncodedContent(new[]
            {
            new KeyValuePair<string, string>("client_id", clientId),
            new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
            new KeyValuePair<string, string>("client_secret", clientSecret),
            new KeyValuePair<string, string>("grant_type", "client_credentials")
        });

            // Send the POST request
            HttpResponseMessage response = await httpClient.PostAsync(uri, requestBody);
            response.EnsureSuccessStatusCode();

            // Parse the response content
            string responseContent = await response.Content.ReadAsStringAsync();
            var jsonResponse = JsonDocument.Parse(responseContent).RootElement;

            // Extract the token and expiration details
            token = jsonResponse.GetProperty("access_token").GetString();
            int expiresIn = jsonResponse.GetProperty("expires_in").GetInt32();

            // Calculate expiration time, minus 3 minutes
            tokenExpiresOn = DateTime.UtcNow.AddSeconds(expiresIn - 180);

            return token;
        }
        private static async Task<JArray> FetchUsersAsync()
        {
            Console.WriteLine("Fetching users...");
            string endpoint = $"{graphEndpoint}/users";
            JArray users = new JArray();

            while (!string.IsNullOrEmpty(endpoint))
            {
                var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                var jsonResponse = JObject.Parse(content);

                // Add current users to the array
                if (jsonResponse["value"] != null)
                {
                    users.Merge(jsonResponse["value"]);
                }

                // Handle pagination
                endpoint = jsonResponse["@odata.nextLink"]?.ToString();
            }

            return users;
        }

        private static async Task BackupUserChatsAsync(JToken user, string rootDirectory, int days)
        {
            string userName = user["displayName"].ToString();
            string userId = user["id"].ToString();

            Console.WriteLine($"Processing user: {userName}");
           // if (user["id"].ToString() == "4e220cac-4b90-4a20-a58b-c3723b90fcab")
           // {

                // Create user directory
                string userDirectory = Path.Combine(rootDirectory, SanitizeFileName(userName));
                Directory.CreateDirectory(userDirectory);

                // Fetch chats for the user
                var chats = await FetchUserChatsAsync(userId);
                Console.WriteLine($"count user chats : {chats.Count}");
                 userChatCount = chats.Count;
                int i = 0;
                foreach (var chat in chats)
                {
                   
                    string chatId = chat["id"].ToString();
                    var messages = await FetchChatMessagesAsync(chatId, days,i);

                    if (messages.Count > 0)
                    {
                        await SaveChatMessagesAsync(chat, messages, userDirectory,i);
                }
                else
                {
                    await SaveEmptyChatMessagesAsync(chat, messages, userDirectory, i);
                }
 Console.WriteLine($"count user sn : {i++}");
                }

            //ww}
        }

        private static async Task<JArray> FetchUserChatsAsync(string userId)
        {
            Console.WriteLine($"Fetching chats for user {userId}...");
            string endpoint = $"{graphEndpoint}/users/{userId}/chats";
            JArray chats = new JArray();

            while (!string.IsNullOrEmpty(endpoint))
            {
                var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                var jsonResponse = JObject.Parse(content);
                
                // Add current chats to the array
                if (jsonResponse["value"] != null)
                {
                    chats.Merge(jsonResponse["value"]);
                }

                // Handle pagination
                endpoint = jsonResponse["@odata.nextLink"]?.ToString();
            }

            return chats;
        }

        private static async Task<JArray> FetchChatMessagesAsync(string chatId, int days, int chatSn)
        {
            Console.WriteLine($"Fetching messages for chat {chatId}...");
            string fromDate = DateTime.UtcNow.AddDays(-days).ToString("o");
            string endpoint = $"{graphEndpoint}/chats/{chatId}/messages?$top=50&$filter=lastModifiedDateTime gt {fromDate}";
            JArray messages = new JArray();

            while (!string.IsNullOrEmpty(endpoint))
            {
                var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var response = await httpClient.SendAsync(request);
                if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    Console.WriteLine($"403 Forbidden: Ensure the app has the required permissions to access chat messages.");
                    Console.WriteLine($"Response: {response}");
                    break;
                }
                Console.WriteLine($"Fetching messages for chat response {response}...");
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                var jsonResponse = JObject.Parse(content);

                // Add messages to the array
                if (jsonResponse["value"] != null)
                {
                    messages.Merge(jsonResponse["value"]);
                }

                // Handle pagination
                endpoint = jsonResponse["@odata.nextLink"]?.ToString();
            }
            Console.WriteLine($"count chat messages : {messages.Count}");
            return messages;
        }

        private static async Task SaveChatMessagesAsync(JToken chat, JArray messages, string userDirectory,int chatSn)
        {
            string chatName = chat["topic"] == null ? "UntitledChat" : chat["topic"].ToString() ?? "Untitled Chat";
            chatName = chatName == "" ? "UnknownTitle" : chatName;
            Console.WriteLine($"Saving messages for chat: {chatName}");

            // Generate chat HTML
            string chatHtml = GenerateChatHtml(chatName, messages, chatSn);

            // Save to file
            string filePath = Path.Combine(userDirectory, $"{SanitizeFileName(chatName)}_{chatSn + 1}.html");
            await File.WriteAllTextAsync(filePath, chatHtml);
        }
        private static async Task SaveEmptyChatMessagesAsync(JToken chat, JArray messages, string userDirectory, int chatSn)
        {
            string chatName = chat["topic"] == null ? "UntitledChat" : chat["topic"].ToString() ?? "Untitled Chat";
            chatName = chatName == "" ? "UnknownTitle" : chatName;
            Console.WriteLine($"Saving messages for chat: {chatName}");

            // Generate chat HTML
            string chatHtml = "No message for this chat";

            // Save to file
            string filePath = Path.Combine(userDirectory, $"{SanitizeFileName(chatName)}_{chatSn + 1}.html");
            await File.WriteAllTextAsync(filePath, chatHtml);
        }

        private static string GenerateChatHtml(string chatName, JArray messages,int chatSn)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<html><head><title>Teams Chat Backup</title></head><body>");
            sb.AppendLine($"<h1>{chatName}</h1> <small>messages on Chat {chatSn + 1 } of {userChatCount}</small>");

            foreach (var message in messages)
            {
                string sender = message["from"] == null ? "NA": message["from"]?.ToString() ?? message["from"]["user"]["displayName"]?.ToString() ?? "Unknown";
                string content = message["body"]["content"]?.ToString() ?? "No Content";
                string timestamp = message["createdDateTime"]?.ToString() ?? "Unknown Date";

                sb.AppendLine($"<div style='border:1 solid blue; margin-bottom:5px'><strong>sender: {sender}</strong> (createdDateTime: {timestamp}): message: {content}</div><hr>");
            }

            sb.AppendLine("</body></html>");
            return sb.ToString();
        }

        private static string SanitizeFileName(string name)
        {
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalidChar, '_');
            }

            return name;
        }
    }
}
