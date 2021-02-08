// <copyright file="DraftNotificationSummary.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;

    /// <summary>
    /// Draft Notification Summary model class.
    /// </summary>
    public class DraftNotificationSummary
    {
        /// <summary>
        /// Gets or sets Notification Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets Updated Date value.
        /// </summary>
        public DateTime LastSavedDate { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether message is to send recursively or not.
        /// </summary>
        public bool IsRecurrence { get; set; }
    }
}