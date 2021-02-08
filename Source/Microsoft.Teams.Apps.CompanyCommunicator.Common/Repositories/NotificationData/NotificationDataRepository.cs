// <copyright file="NotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Reflection;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Repository of the notification data in the table storage.
    /// </summary>
    public class NotificationDataRepository : BaseRepository<NotificationDataEntity>
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private readonly IConfiguration configuration;
        private readonly string graphQuery = $"https://graph.microsoft.com/v1.0/$batch";
        private readonly string[] scopesVal = { "https://graph.microsoft.com/.default" };
        private readonly ADGroupsDataRepository adGroupsDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Represents the application configuration.</param>
        /// <param name="adGroupsDataRepository">Represents ADGroupsDataRepository object. </param>
        /// <param name="tableRowKeyGenerator">Table row key generator service.</param>
        /// <param name="isFromAzureFunction">Flag to show if created from Azure Function.</param>
        public NotificationDataRepository(
            IConfiguration configuration,
            ADGroupsDataRepository adGroupsDataRepository,
            TableRowKeyGenerator tableRowKeyGenerator,
            bool isFromAzureFunction = false)
            : base(
                configuration,
                PartitionKeyNames.NotificationDataTable.TableName,
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                isFromAzureFunction)
        {
            this.adGroupsDataRepository = adGroupsDataRepository;
            this.TableRowKeyGenerator = tableRowKeyGenerator;
            this.configuration = configuration;
        }

        /// <summary>
        /// Gets table row key generator.
        /// </summary>
        public TableRowKeyGenerator TableRowKeyGenerator { get; }

        /// <summary>
        /// Get all draft notification entities from the table storage.
        /// </summary>
        /// <returns>All draft notification entities.</returns>
        public async Task<IEnumerable<NotificationDataEntity>> GetAllDraftNotificationsAsync()
        {
            var result = await this.GetAllAsync(PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition);

            return result;
        }

        /// <summary>
        /// Get all sent notification entities from the table storage.
        /// </summary>
        /// <param name="notificationId"> Sent Notification id.</param>
        /// <returns>All sent notification entities.</returns>
        public async Task<NotificationDataEntity> GetSentNotificationAsync(string notificationId)
        {
            var result = await this.GetAsync(PartitionKeyNames.NotificationDataTable.SentNotificationsPartition, notificationId);

            return result;
        }

        /// <summary>
        /// Get all sent notification entities from the table storage.
        /// </summary>
        /// <param name="notificationId"> Sent Notification id.</param>
        /// <returns>All sent notification entities.</returns>
        public async Task<NotificationDataEntity> GetScheduleSentNotificationAsync(string notificationId)
        {
            var result = await this.GetAsync(PartitionKeyNames.NotificationDataTable.ScheduleSentNotificationsPartition, notificationId);

            return result;
        }

        /// <summary>
        /// Get the top 25 most recently sent notification entities from the table storage.
        /// </summary>
        /// <returns>The top 25 most recently sent notification entities.</returns>
        public async Task<IEnumerable<NotificationDataEntity>> GetMostRecentSentNotificationsAsync()
        {
            var result = await this.GetAllAsync(PartitionKeyNames.NotificationDataTable.SentNotificationsPartition, 25);

            return result;
        }

        /// <summary>
        /// Move a draft notification from draft to sent partition.
        /// </summary>
        /// <param name="draftNotificationEntity">The draft notification instance to be moved to the sent partition.</param>
        /// <param name="isScheduleorREcurrence">Indicates whether message is schedule or recurrence.</param>
        /// <returns>The new SentNotification ID.</returns>
        public async Task<string> MoveDraftToSentPartitionAsync(NotificationDataEntity draftNotificationEntity, bool isScheduleorREcurrence)
        {
            if (draftNotificationEntity == null)
            {
                return string.Empty;
            }

            var newId = this.TableRowKeyGenerator.CreateNewKeyOrderingMostRecentToOldest();

            // Create a sent notification based on the draft notification.
            var sentNotificationEntity = new NotificationDataEntity
            {
                PartitionKey = isScheduleorREcurrence ? PartitionKeyNames.NotificationDataTable.ScheduleSentNotificationsPartition : PartitionKeyNames.NotificationDataTable.SentNotificationsPartition,
                RowKey = newId,
                Id = newId,
                Title = draftNotificationEntity.Title,
                ImageLink = draftNotificationEntity.ImageLink,
                Summary = draftNotificationEntity.Summary,
                Author = draftNotificationEntity.Author,
                ButtonTitle = draftNotificationEntity.ButtonTitle,
                ButtonLink = draftNotificationEntity.ButtonLink,
                ButtonTitle2 = draftNotificationEntity.ButtonTitle2,
                ButtonLink2 = draftNotificationEntity.ButtonLink2,
                CreatedBy = draftNotificationEntity.CreatedBy,
                CreatedDate = draftNotificationEntity.CreatedDate,
                SentDate = null,
                IsDraft = false,
                Teams = draftNotificationEntity.Teams,
                Rosters = draftNotificationEntity.Rosters,
                AllUsers = draftNotificationEntity.AllUsers,
                ADGroups = draftNotificationEntity.ADGroups,
                MessageVersion = draftNotificationEntity.MessageVersion,
                Succeeded = 0,
                Failed = 0,
                Throttled = 0,
                TotalMessageCount = draftNotificationEntity.TotalMessageCount,
                IsCompleted = false,
                SendingStartedDate = DateTime.UtcNow,
                IsScheduled = draftNotificationEntity.IsScheduled,
                ScheduleDate = draftNotificationEntity.ScheduleDate,
                IsRecurrence = draftNotificationEntity.IsRecurrence,
                Repeats = draftNotificationEntity.Repeats,
                RepeatFor = Convert.ToInt32(draftNotificationEntity.RepeatFor),
                RepeatFrequency = draftNotificationEntity.RepeatFrequency,
                WeekSelection = draftNotificationEntity.WeekSelection,
                RepeatStartDate = draftNotificationEntity.RepeatStartDate,
                RepeatEndDate = draftNotificationEntity.RepeatEndDate,
            };
            await this.CreateOrUpdateAsync(sentNotificationEntity);

            // Delete the draft notification.
            await this.DeleteAsync(draftNotificationEntity);

            return newId;
        }

        /// <summary>
        /// Move a draft notification from draft to sent partition.
        /// </summary>
        /// <param name="masterNotificationEntity">The draft notification instance to be moved to the sent partition.</param>
        /// <returns>The new SentNotification ID.</returns>
        public async Task<string> CopyToSentPartitionAsync(NotificationDataEntity masterNotificationEntity)
        {
            if (masterNotificationEntity == null)
            {
                return string.Empty;
            }

            var newId = this.TableRowKeyGenerator.CreateNewKeyOrderingMostRecentToOldest();

            // Create a sent notification based on the master notification.
            var sentNotificationEntity = new NotificationDataEntity
            {
                PartitionKey = PartitionKeyNames.NotificationDataTable.SentNotificationsPartition,
                RowKey = newId,
                Id = newId,
                Title = masterNotificationEntity.Title,
                ImageLink = masterNotificationEntity.ImageLink,
                Summary = masterNotificationEntity.Summary,
                Author = masterNotificationEntity.Author,
                ButtonTitle = masterNotificationEntity.ButtonTitle,
                ButtonLink = masterNotificationEntity.ButtonLink,
                ButtonTitle2 = masterNotificationEntity.ButtonTitle2,
                ButtonLink2 = masterNotificationEntity.ButtonLink2,
                CreatedBy = masterNotificationEntity.CreatedBy,
                CreatedDate = masterNotificationEntity.CreatedDate,
                SentDate = null,
                IsDraft = false,
                Teams = masterNotificationEntity.Teams,
                Rosters = masterNotificationEntity.Rosters,
                AllUsers = masterNotificationEntity.AllUsers,
                ADGroups = masterNotificationEntity.ADGroups,
                MessageVersion = masterNotificationEntity.MessageVersion,
                Succeeded = 0,
                Failed = 0,
                Throttled = 0,
                TotalMessageCount = masterNotificationEntity.TotalMessageCount,
                IsCompleted = false,
                SendingStartedDate = DateTime.UtcNow,
                IsScheduled = masterNotificationEntity.IsScheduled,
                ScheduleDate = masterNotificationEntity.ScheduleDate,
                IsRecurrence = masterNotificationEntity.IsRecurrence,
                Repeats = masterNotificationEntity.Repeats,
                RepeatFor = Convert.ToInt32(masterNotificationEntity.RepeatFor),
                RepeatFrequency = masterNotificationEntity.RepeatFrequency,
                WeekSelection = masterNotificationEntity.WeekSelection,
                RepeatStartDate = masterNotificationEntity.RepeatStartDate,
                RepeatEndDate = masterNotificationEntity.RepeatEndDate,
            };
            await this.CreateOrUpdateAsync(sentNotificationEntity);

            return newId;
        }

        /// <summary>
        /// Duplicate an existing draft notification.
        /// </summary>
        /// <param name="notificationEntity">The notification entity to be duplicated.</param>
        /// <param name="createdBy">Created by.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task DuplicateDraftNotificationAsync(
            NotificationDataEntity notificationEntity,
            string createdBy)
        {
            var newId = this.TableRowKeyGenerator.CreateNewKeyOrderingOldestToMostRecent();

            var newNotificationEntity = new NotificationDataEntity
            {
                PartitionKey = PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                RowKey = newId,
                Id = newId,
                Title = notificationEntity.Title + " (copy)",
                ImageLink = notificationEntity.ImageLink,
                Summary = notificationEntity.Summary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                CreatedBy = createdBy,
                CreatedDate = DateTime.UtcNow,
                IsDraft = true,
                Teams = notificationEntity.Teams,
                Rosters = notificationEntity.Rosters,
                AllUsers = notificationEntity.AllUsers,
                ADGroups = notificationEntity.ADGroups,
                IsScheduled = notificationEntity.IsScheduled,
                ScheduleDate = notificationEntity.ScheduleDate,
                IsRecurrence = notificationEntity.IsRecurrence,
                Repeats = notificationEntity.Repeats,
                RepeatFor = Convert.ToInt32(notificationEntity.RepeatFor),
                RepeatFrequency = notificationEntity.RepeatFrequency,
                WeekSelection = notificationEntity.WeekSelection,
                RepeatStartDate = notificationEntity.RepeatStartDate,
                RepeatEndDate = notificationEntity.RepeatEndDate,
                ButtonTitle2 = notificationEntity.ButtonTitle2,
                ButtonLink2 = notificationEntity.ButtonLink2,
            };

            await this.CreateOrUpdateAsync(newNotificationEntity);
        }
    }
}
