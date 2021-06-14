
namespace Icebreaker.Helpers
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Azure.Documents;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    /// <summary>
    /// Represents a user that was matched with completed meetup counts
    /// </summary>
    public class UserMatchInfoWithCount : UserMatchInfo
    {
        /// <summary>
        /// Gets or sets the meetup count for user
        /// </summary>
        [JsonProperty("meetupCount")]
        public int MeetupCount { get; set; }

        /// <summary>
        /// Gets or sets the user display image url
        /// </summary>
        [JsonProperty("userDisplayUrl")]
        public Uri UserDisplayUrl { get; set; }

        /// <summary>
        /// Gets or sets the user display name
        /// </summary>
        [JsonProperty("userDisplayName")]
        public string UserDisplayName { get; set; }
    }

    public class UserMessagesResponse
    {
        [JsonProperty("value")]
        public List<Message> Messages { get; set; }
    }
}