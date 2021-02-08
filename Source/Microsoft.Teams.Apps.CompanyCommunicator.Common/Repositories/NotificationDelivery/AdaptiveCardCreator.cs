// <copyright file="AdaptiveCardCreator.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.NotificationDelivery
{
    using System;
    using AdaptiveCards;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Adaptive Card Creator service.
    /// </summary>
    public class AdaptiveCardCreator
    {
        /// <summary>
        /// Acknowledgment button text.
        /// </summary>
        public static readonly string Acknowledgment = "Acknowledge";

        /// <summary>
        /// Creates an adaptive card.
        /// </summary>
        /// <param name="notificationDataEntity">Notification data entity.</param>
        /// <param name="notificationId">notification ID value.</param>
        /// <returns>An adaptive card.</returns>
        public AdaptiveCard CreateAdaptiveCard(NotificationDataEntity notificationDataEntity, string notificationId)
        {
            return this.CreateAdaptiveCard(
                notificationDataEntity.Title,
                notificationDataEntity.ImageLink,
                notificationDataEntity.Summary,
                notificationDataEntity.Author,
                notificationDataEntity.ButtonTitle,
                notificationDataEntity.ButtonLink,
                notificationDataEntity.ButtonTitle2,
                notificationDataEntity.ButtonLink2,
                notificationId);
        }

        /// <summary>
        /// Create an adaptive card instance.
        /// </summary>
        /// <param name="title">The adaptive card's title value.</param>
        /// <param name="imageUrl">The adaptive card's image URL.</param>
        /// <param name="summary">The adaptive card's summary value.</param>
        /// <param name="author">The adaptive card's author value.</param>
        /// <param name="buttonTitle">The adaptive card's button title value.</param>
        /// <param name="buttonUrl">The adaptive card's button URL value.</param>
        /// <param name="buttonTitle2">The adaptive card's second button title value.</param>
        /// <param name="buttonUrl2">The adaptive card's second button URL value.</param>
        /// <param name="notificationid">The notification id.</param>
        /// <param name="isAcknowledgementRequired">Indicates whether acknowledge button is required or not.</param>
        /// <returns>The created adaptive card instance.</returns>
        public AdaptiveCard CreateAdaptiveCard(
            string title,
            string imageUrl,
            string summary,
            string author,
            string buttonTitle,
            string buttonUrl,
            string buttonTitle2,
            string buttonUrl2,
            string notificationid,
            bool isAcknowledgementRequired = true)
        {
            var version = new AdaptiveSchemaVersion(1, 0);
            AdaptiveCard card = new AdaptiveCard(version);

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = title,
                Size = AdaptiveTextSize.ExtraLarge,
                Weight = AdaptiveTextWeight.Bolder,
                Wrap = true,
            });

            if (!string.IsNullOrWhiteSpace(imageUrl) && imageUrl != "https://")
            {
                card.Body.Add(new AdaptiveImage()
                {
                    Url = new Uri(imageUrl, UriKind.RelativeOrAbsolute),
                    Spacing = AdaptiveSpacing.Default,
                    Size = AdaptiveImageSize.Stretch,
                    AltText = string.Empty,
                });
            }

            if (!string.IsNullOrWhiteSpace(summary))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = summary,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(author))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = author,
                    Size = AdaptiveTextSize.Small,
                    Weight = AdaptiveTextWeight.Lighter,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle)
                && !string.IsNullOrWhiteSpace(buttonUrl) && buttonUrl != "https://")
            {
                card.Actions.Add(new AdaptiveOpenUrlAction()
                {
                    Title = buttonTitle,
                    Url = new Uri(buttonUrl, UriKind.RelativeOrAbsolute),
                });
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle2)
                && !string.IsNullOrWhiteSpace(buttonUrl2) && buttonUrl2 != "https://")
            {
                card.Actions.Add(new AdaptiveOpenUrlAction()
                {
                    Title = buttonTitle2,
                    Url = new Uri(buttonUrl2, UriKind.RelativeOrAbsolute),
                });
            }

            // Acknowledgment button.
            if (isAcknowledgementRequired)
            {
                card.Actions.Add(new AdaptiveSubmitAction
                {
                    Title = Acknowledgment,
                    Data = new MessageCard { Title = title, ImageUrl = imageUrl, Summary = summary, Author = author, ButtonTitle = buttonTitle, ButtonUrl = buttonUrl, ButtonTitle2 = buttonTitle2, ButtonUrl2 = buttonUrl2, NotificationId = notificationid },
                });
            }

            return card;
        }
    }
}
