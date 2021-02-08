// <copyright file="ADGroup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
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
    }
}
