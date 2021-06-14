namespace Icebreaker.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure;
    using Microsoft.Graph;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
    using Newtonsoft.Json;

    public class GraphHelper
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphHelper"/> class.
        /// </summary>
        /// <param name="telemetryClient">logging telemetry</param>
        public GraphHelper(TelemetryClient telemetryClient)
        {
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Returns the user object for AadObjectId
        /// </summary>
        /// <param name="userId">user id for user</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<User> GetUser(string userId)
        {
            this.telemetryClient.TrackTrace("Sending request to fetch the user object.", SeverityLevel.Information);
            var userObj = await this.SendGraphRequest<User>($"https://graph.microsoft.com/v1.0/users/{userId}", HttpMethod.Get);

            return userObj;
        }

        /// <summary>
        /// Returns the messages fetched for User
        /// </summary>
        /// <param name="userId">user id for user</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<List<Message>> GetUserMessages(string userId)
        {
            this.telemetryClient.TrackTrace("Sending request to fetch the user messages.", SeverityLevel.Information);
            var userMessages = await this.SendGraphRequest<UserMessagesResponse>($"https://graph.microsoft.com/v1.0/users/{userId}/messages?$filter=isDraft eq false", HttpMethod.Get);

            return userMessages.Messages;
        }

        /// <summary>
        /// Returns a base64 image url for user
        /// </summary>
        /// <param name="user">User object</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<Uri> GetUserDisplay(User user)
        {
            try
            {
                this.telemetryClient.TrackTrace("Sending request to fetch the user display picture.", SeverityLevel.Information);
                var bytes = await GetStreamWithAuthAsync($"https://graph.microsoft.com/v1.0/users/{user.Id}/photos/48X48/$value");

                if (bytes != null)
                {
                    var pictureUrl = "data:image/png;base64," + Convert.ToBase64String(bytes);
                    return new Uri(pictureUrl);
                }
            }
            catch
            {
                this.telemetryClient.TrackTrace("Exception encountered!", SeverityLevel.Information);
            }

            return new Uri($"https://ui-avatars.com/api/?rounded=true&name={user.GivenName}+{user.Surname}&background=cfe0d6&color=154229&bold=true");
        }

        private static async Task<string> GetToken()
        {
            var re = new Regex(@"https://(.+?).sharepoint.com");
            var tenantId = CloudConfigurationManager.GetSetting("TenantId") ?? $"{re.Match(CloudConfigurationManager.GetSetting("RootSiteUrl").ToLower()).Groups[1].Value}.onmicrosoft.com";

            var authContext = new AuthenticationContext($"https://login.microsoftonline.com/{tenantId}");
            var cc = new ClientCredential(CloudConfigurationManager.GetSetting("MicrosoftAppId"), CloudConfigurationManager.GetSetting("MicrosoftAppPassword"));

            var result = await authContext.AcquireTokenAsync("https://graph.microsoft.com", cc);
            return result.AccessToken;
        }

        private static async Task<byte[]> GetStreamWithAuthAsync(string endpoint)
        {
            var token = await GetToken();
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            client.DefaultRequestHeaders.Add("Accept", "application/json");

            using (var response = await client.GetAsync(endpoint))
            {
                if (response.IsSuccessStatusCode)
                {
                    var stream = await response.Content.ReadAsStreamAsync();
                    byte[] bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    return bytes;
                }
                else
                {
                    return null;
                }
            }
        }

        private async Task<HttpResponseMessage> SendGraphRequest(string apiUrl, HttpMethod method, object bodyData = null, int retryCount = 3, int delay = 500)
        {
            int retryAttempts = 0;
            int backoffInterval = delay;
            var lastResult = new HttpResponseMessage();

            while (retryAttempts < retryCount)
            {
                try
                {
                    var token = await GetToken();

                    using (var request = new HttpRequestMessage(method, apiUrl))
                    {
                        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                        if (bodyData != null)
                        {
                            var bodyContent = bodyData is string ? bodyData.ToString() : JsonConvert.SerializeObject(bodyData);
                            request.Content = new StringContent(bodyContent, Encoding.UTF8, "application/json");
                        }

                        var result = await HttpClient.SendAsync(request);
                        lastResult = result;

                        if (result.IsSuccessStatusCode)
                        {
                            return result;
                        }
                        else
                        {
                            this.telemetryClient.TrackTrace($"Error reason phrase: {result.ReasonPhrase}", SeverityLevel.Information);
                            var response = result.Content.ReadAsStringAsync();
                            this.telemetryClient.TrackTrace($"Response: {response.Result}", SeverityLevel.Information);
                            throw new Exception();
                        }
                    }
                }
                catch
                {
                    this.telemetryClient.TrackTrace("Retrying...", SeverityLevel.Information);
                    retryAttempts++;
                    if (retryAttempts == retryCount)
                    {
                        throw;
                    }
                    else
                    {
                        await Task.Delay(backoffInterval);
                        backoffInterval = backoffInterval * 2;
                    }
                }
            }

            return lastResult;
        }

        private async Task<T> SendGraphRequest<T>(string apiUrl, HttpMethod method, object bodyData = null)
        {
            var response = await this.SendGraphRequest(apiUrl, method, bodyData);
            if (response.IsSuccessStatusCode)
            {
                this.telemetryClient.TrackTrace("Request executed successfully", SeverityLevel.Information);
                string data = await response.Content.ReadAsStringAsync();
                T obj = JsonConvert.DeserializeObject<T>(data);
                return obj;
            }
            else
            {
                return default(T);
            }
        }
    }
}