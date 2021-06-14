
namespace Icebreaker.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web.Hosting;
    using System.Web.Http;
    using Icebreaker.Helpers;
    using Icebreaker.Interfaces;
    using Icebreaker.Properties;
    using Microsoft.ApplicationInsights;
    using Microsoft.Azure;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    /// <summary>
    /// API controller to get leaders data.
    /// </summary>
    public class UserMeetupInfoController : ApiController
    {
        private const string KeyHeaderName = "X-Key";
        private readonly MicrosoftAppCredentials botCredentials;
        private readonly string apiKey;
        private readonly string botDisplayName;
        private readonly IUserMeetupInfoService userMeetupInfoService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserMeetupInfoController"/> class.
        /// </summary>
        /// <param name="userMeetupInfoService">User Meetup Service</param>
        /// <param name="botCredentials">The bot AAD credentials</param>
        /// <param name="secretsHelper">Secrets helper to fetch secrets</param>
        public UserMeetupInfoController(IUserMeetupInfoService userMeetupInfoService, MicrosoftAppCredentials botCredentials, ISecretsHelper secretsHelper)
        {
            this.botCredentials = botCredentials;
            this.apiKey = secretsHelper.Key;
            this.botDisplayName = CloudConfigurationManager.GetSetting("BotDisplayName");
            this.userMeetupInfoService = userMeetupInfoService;
        }

        /// <summary>
        /// Action to process matches
        /// </summary>
        /// <returns>Success (1) or failure (-1) code</returns>
        [Route("api/getleaders")]
        public async Task<HttpResponseMessage> GetAsync()
        {
            IEnumerable<string> keys;
            if (this.Request.Headers.TryGetValues(KeyHeaderName, out keys))
            {
                var isKeyMatch = keys.Any() && object.Equals(keys.First(), this.apiKey);
                if (isKeyMatch)
                {
                    // Get the token here to proactively trigger a refresh if the cached token is expired
                    // This avoids a race condition in MicrosoftAppCredentials.GetTokenAsync that can lead it to return an expired token
                    await this.botCredentials.GetTokenAsync();

                    var allUserMeetupLeaders = await this.userMeetupInfoService.GetUserMeetupLeaders();

                    return new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent(JsonConvert.SerializeObject(allUserMeetupLeaders), Encoding.UTF8, "application/json"),
                    };
                }
            }

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(string.Empty, Encoding.UTF8, "application/json"),
            };
        }
    }
}