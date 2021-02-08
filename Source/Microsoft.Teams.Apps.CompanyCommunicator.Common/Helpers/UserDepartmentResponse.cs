// <copyright file="UserDepartmentResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers
{
    using Newtonsoft.Json;

    /// <summary>
    /// User Department Response class.
    /// </summary>
    public class UserDepartmentResponse
    {
        /// <summary>
        /// Gets or sets id.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets department.
        /// </summary>
        [JsonProperty("department")]
        public string Department { get; set; }
    }
}
