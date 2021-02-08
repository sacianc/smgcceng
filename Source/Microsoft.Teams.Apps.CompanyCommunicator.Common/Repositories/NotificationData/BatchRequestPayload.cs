// <copyright file="BatchRequestPayload.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System.Collections.Generic;

    /// <summary>
    /// creating <see cref="BatchRequestPayload"/> class.
    /// </summary>
    public class BatchRequestPayload
    {
        /// <summary>
        /// Gets or sets this method to get list of requests of each request in batch call.
        /// </summary>
#pragma warning disable CA2227 // Collection properties should be read only
        public List<dynamic> Requests { get; set; }
#pragma warning restore CA2227 // Collection properties should be read only
    }
}