// <copyright file="ADGroup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// This Model is for AD Group.
    /// </summary>
    public class ADGroup
    {
        /// <summary>
        /// Gets or sets this method to get ID from AAD.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets this method for displayName from AAD.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets this method for mail from AAD.
        /// </summary>
        [JsonProperty("mail")]
        public string Mail { get; set; }

        /// <summary>
        /// Gets or sets this method for mailNickname from AAD.
        /// </summary>
        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        /// <summary>
        /// Gets or sets this method for User Principal name from AAD.
        /// </summary>
        [JsonProperty("upn")]
        public string UPN { get; set; }
    }
}
