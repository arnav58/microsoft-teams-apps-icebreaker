
namespace Icebreaker.Helpers
{
    using System.Collections.Generic;
    using Microsoft.Azure.Documents;
    using Newtonsoft.Json;

    /// <summary>
    /// Represents a user that was matched
    /// </summary>
    public class UserMatchInfo : Document
    {
        /// <summary>
        /// Gets or sets the user's id in Teams (29:xxx).
        /// This is also the <see cref="Resource.Id"/>.
        /// </summary>
        [JsonIgnore]
        public string UserId
        {
            get { return this.Id; }
            set { this.Id = value; }
        }

        /// <summary>
        /// Gets or sets the tenant id
        /// </summary>
        [JsonProperty("tenantId")]
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the user aad object id
        /// </summary>
        [JsonProperty("userAadObjectId")]
        public string UserAadObjectId { get; set; }

        /// <summary>
        /// Gets or sets the user principal name
        /// </summary>
        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets the role
        /// </summary>
        [JsonProperty("role")]
        public string Role { get; set; }
    }
}