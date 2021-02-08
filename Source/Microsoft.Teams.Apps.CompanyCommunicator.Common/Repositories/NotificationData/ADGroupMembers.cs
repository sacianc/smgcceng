// <copyright file="ADGroupMembers.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using Newtonsoft.Json;

    /// <summary>
    /// User data model class
    /// </summary>
    public class ADGroupMembers
    {
        /// <summary>
        /// Gets or sets the ID from AAD for a particular user.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets display name of the user
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets user principal name of the user
        /// </summary>
        [JsonProperty("userPrincipalName")]
        public string UserPrincipleName { get; set; }

        /// <summary>
        /// Gets or sets department of the user
        /// </summary>
        [JsonProperty("@odata.type")]
        public string Type { get; set; }
    }
}
