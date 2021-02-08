// <copyright file="MessageCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// This Model is for Acknowledgment message card.
    /// </summary>
    public class MessageCard
    {
        /// <summary>
        /// Gets or sets the title of communication message.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the image URL of communication message.
        /// </summary>
        public string ImageUrl { get; set; }

        /// <summary>
        /// Gets or sets the full summary of communication message-
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the author of communication message.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title of communication message.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button URL of communication message.
        /// </summary>
        public string ButtonUrl { get; set; }

        /// <summary>
        /// Gets or sets the second Button Title of communication message.
        /// </summary>
        public string ButtonTitle2 { get; set; }

        /// <summary>
        /// Gets or sets the second Button URL of communication message.
        /// </summary>
        public string ButtonUrl2 { get; set; }

        /// <summary>
        /// Gets or sets this method to get notificationId from database.
        /// </summary>
        [JsonProperty("notificationId")]
        public string NotificationId { get; set; }
    }
}
