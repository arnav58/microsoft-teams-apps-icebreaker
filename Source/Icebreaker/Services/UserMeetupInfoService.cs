
namespace Icebreaker.Services
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Icebreaker.Helpers;
    using Icebreaker.Helpers.AdaptiveCards;
    using Icebreaker.Interfaces;
    using Icebreaker.Properties;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Graph;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Implements the logic to get the leaders
    /// </summary>
    public class UserMeetupInfoService : IUserMeetupInfoService
    {
        private readonly IBotDataProvider dataProvider;
        private readonly ConversationHelper conversationHelper;
        private readonly TelemetryClient telemetryClient;
        private readonly BotAdapter botAdapter;
        private readonly GraphHelper graphHelper;
        private readonly string botDisplayName;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserMeetupInfoService"/> class.
        /// </summary>
        /// <param name="dataProvider">The data provider to use</param>
        /// <param name="conversationHelper">Conversation helper instance to notify team members</param>
        /// <param name="telemetryClient">The telemetry client to use</param>
        /// <param name="botAdapter">Bot adapter.</param>
        public UserMeetupInfoService(IBotDataProvider dataProvider, ConversationHelper conversationHelper, TelemetryClient telemetryClient, BotAdapter botAdapter)
        {
            this.dataProvider = dataProvider;
            this.conversationHelper = conversationHelper;
            this.telemetryClient = telemetryClient;
            this.botAdapter = botAdapter;
            this.graphHelper = new GraphHelper(telemetryClient);
            this.botDisplayName = CloudConfigurationManager.GetSetting("BotDisplayName");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserMeetupInfoService"/> class.
        /// </summary>
        /// <returns>Leaders Info</returns>
        public async Task<IList<UserMatchInfoWithCount>> GetUserMeetupLeaders()
        {
            var meetingContent = string.Format(Resources.MeetupContent, this.botDisplayName);

            var usersMatchedByApp = await this.dataProvider.GetUsersMatchedByApp();

            foreach (UserMatchInfoWithCount userMatched in usersMatchedByApp)
            {
                var messages = await this.graphHelper.GetUserMessages(userMatched.UserAadObjectId);
                var filtedMessages = messages.FindAll(m => m.BodyPreview.Contains(meetingContent) && !m.Flag.FlagStatus.Equals(FollowupFlagStatus.Flagged));

                User user = await this.graphHelper.GetUser(userMatched.UserAadObjectId);
                Uri userDisplay = await this.graphHelper.GetUserDisplay(user);

                userMatched.MeetupCount = filtedMessages.Count;
                userMatched.UserDisplayUrl = userDisplay;
                userMatched.UserDisplayName = user.DisplayName;
            }

            return usersMatchedByApp.OrderByDescending(user => user.MeetupCount).ToList();
        }
    }
}