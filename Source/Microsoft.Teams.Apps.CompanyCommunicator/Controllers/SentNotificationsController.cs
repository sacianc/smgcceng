// <copyright file="SentNotificationsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ScheduleNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.NotificationDelivery;

    /// <summary>
    /// Controller for the sent notification data.
    /// </summary>
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/sentNotifications")]
    public class SentNotificationsController : ControllerBase
    {
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly SentNotificationDataRepository sentNotificationDataRepository;
        private readonly NotificationDelivery notificationDelivery;
        private readonly TeamDataRepository teamDataRepository;
        private readonly ScheduleNotificationDataRepository scheduleNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="SentNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository service that deals with the table storage in azure.</param>
        /// <param name="notificationDelivery">Notification delivery service instance.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        /// <param name="scheduleNotificationDataRepository">Schedule notification data repository.</param>
        /// <param name="sentNotificationDataRepository">Sent Notification Data Repository.</param>
        public SentNotificationsController(
            NotificationDataRepository notificationDataRepository,
            NotificationDelivery notificationDelivery,
            TeamDataRepository teamDataRepository,
            ScheduleNotificationDataRepository scheduleNotificationDataRepository,
            SentNotificationDataRepository sentNotificationDataRepository)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.notificationDelivery = notificationDelivery;
            this.teamDataRepository = teamDataRepository;
            this.scheduleNotificationDataRepository = scheduleNotificationDataRepository;
            this.sentNotificationDataRepository = sentNotificationDataRepository;
        }

        /// <summary>
        /// Send a notification, which turns a draft to be a sent notification.
        /// </summary>
        /// <param name="draftNotification">An instance of <see cref="DraftNotification"/> class.</param>
        /// <returns>The result of an action method.</returns>
        [HttpPost]
        public async Task<IActionResult> CreateSentNotificationAsync([FromBody]DraftNotification draftNotification)
        {
            var draftNotificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                draftNotification.Id);
            if (draftNotificationEntity == null)
            {
                return this.NotFound();
            }

            await this.notificationDelivery.SendAsync(draftNotificationEntity);

            return this.Ok();
        }

        /// <summary>
        /// Get most recently sent notification summaries.
        /// </summary>
        /// <returns>A list of <see cref="SentNotificationSummary"/> instances.</returns>
        [HttpGet]
        public async Task<IEnumerable<SentNotificationSummary>> GetSentNotificationsAsync()
        {
            var notificationEntities = await this.notificationDataRepository.GetMostRecentSentNotificationsAsync();

            var result = new List<SentNotificationSummary>();
            var rowKeysFilter = string.Empty;

            foreach (var notificationEntity in notificationEntities)
            {
                var singleRowKeyFilter = TableQuery.GenerateFilterCondition(
                    nameof(TableEntity.PartitionKey),
                    QueryComparisons.Equal,
                    notificationEntity.Id);

                if (string.IsNullOrWhiteSpace(rowKeysFilter))
                {
                    rowKeysFilter = singleRowKeyFilter;
                }
                else
                {
                    rowKeysFilter = TableQuery.CombineFilters(rowKeysFilter, TableOperators.Or, singleRowKeyFilter);
                }
            }

            var sentNotificationLogs = await this.sentNotificationDataRepository.GetCustomWithFilterAsync(rowKeysFilter);

            foreach (var notificationEntity in notificationEntities)
            {
                int acknowledgementCount = sentNotificationLogs.ToList().Where(o => o.PartitionKey == notificationEntity.Id).Count();
                var summary = new SentNotificationSummary
                {
                    Id = notificationEntity.Id,
                    Title = notificationEntity.Title,
                    CreatedDateTime = notificationEntity.CreatedDate,
                    SentDate = notificationEntity.SentDate,
                    Succeeded = notificationEntity.Succeeded,
                    Failed = notificationEntity.Failed,
                    Throttled = notificationEntity.Throttled,
                    Acknowledged = acknowledgementCount,
                    TotalMessageCount = notificationEntity.TotalMessageCount,
                    IsCompleted = notificationEntity.IsCompleted,
                    SendingStartedDate = notificationEntity.SendingStartedDate,
                    IsRecurrence = notificationEntity.IsRecurrence,
                };

                result.Add(summary);
            }

            return result;
        }

        /// <summary>
        /// Get scheduled notification summaries.
        /// </summary>
        /// <returns>A list of <see cref="SentNotificationSummary"/> instances.</returns>
        [HttpGet("scheduled")]
        public async Task<IEnumerable<SentNotificationSummary>> GetScheduledNotificationsAsync()
        {
            var scheduleNotificationEntities = await this.scheduleNotificationDataRepository.GetScheduledNotificationsAsync();
            var result = new List<SentNotificationSummary>();

            foreach (var scheduleNotificationEntity in scheduleNotificationEntities)
            {
                var notificationEntity = await this.notificationDataRepository.GetScheduleSentNotificationAsync(scheduleNotificationEntity.NotificationId);
                if (notificationEntity != null)
                {
                    var summary = new SentNotificationSummary
                    {
                        Id = notificationEntity.Id,
                        Title = notificationEntity.Title,
                        CreatedDateTime = notificationEntity.CreatedDate,
                        SentDate = scheduleNotificationEntity.NotificationDate,
                        Succeeded = notificationEntity.Succeeded,
                        Failed = notificationEntity.Failed,
                        Throttled = notificationEntity.Throttled,
                        TotalMessageCount = notificationEntity.TotalMessageCount,
                        IsCompleted = notificationEntity.IsCompleted,
                        SendingStartedDate = notificationEntity.SendingStartedDate,
                        IsRecurrence = notificationEntity.IsRecurrence,
                    };

                    result.Add(summary);
                }
            }

            return result;
        }

        /// <summary>
        /// Get a sent notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested sent notification.</param>
        /// <returns>Required sent notification.</returns>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetSentNotificationByIdAsync(string id)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.SentNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var rowKeysFilter = string.Empty;

            var singleRowKeyFilter = TableQuery.GenerateFilterCondition(
                nameof(TableEntity.PartitionKey),
                QueryComparisons.Equal,
                notificationEntity.Id);

            var sentNotificationLogs = await this.sentNotificationDataRepository.GetCustomWithFilterAsync(singleRowKeyFilter);

            var result = new SentNotification
            {
                Id = notificationEntity.Id,
                Title = notificationEntity.Title,
                ImageLink = notificationEntity.ImageLink,
                Summary = notificationEntity.Summary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                ButtonTitle2 = notificationEntity.ButtonTitle2,
                ButtonLink2 = notificationEntity.ButtonLink2,
                CreatedDateTime = notificationEntity.CreatedDate,
                IsRecurrence = notificationEntity.IsRecurrence,
                SentDate = notificationEntity.SentDate,
                Succeeded = notificationEntity.Succeeded,
                Failed = notificationEntity.Failed,
                Throttled = notificationEntity.Throttled,
                Acknowledged = sentNotificationLogs.Count(),
                TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Teams),
                RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Rosters),
                AllUsers = notificationEntity.AllUsers,
                SendingStartedDate = notificationEntity.SendingStartedDate,
                Repeats = notificationEntity.Repeats,
                RepeatStartDate = notificationEntity.RepeatStartDate,
                RepeatFor = notificationEntity.RepeatFor,
                RepeatFrequency = notificationEntity.RepeatFrequency,
                WeekSelection = notificationEntity.WeekSelection,
                RepeatEndDate = notificationEntity.RepeatEndDate,
            };

            return this.Ok(result);
        }

        /// <summary>
        /// Get a schedule notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested schedule notification.</param>
        /// <returns>Required schedule notification.</returns>
        [HttpGet("schedule/{id}")]
        public async Task<IActionResult> GetScheduleNotificationByIdAsync(string id)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.ScheduleSentNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var scheduleNotificationEntity = await this.scheduleNotificationDataRepository.GetAsync(
                PartitionKeyNames.ScheduleNotificationDataTable.ScheduleNotificationsPartition,
                id);
            if (scheduleNotificationEntity == null)
            {
                return this.NotFound();
            }

            var result = new SentNotification
            {
                Id = notificationEntity.Id,
                Title = notificationEntity.Title,
                ImageLink = notificationEntity.ImageLink,
                Summary = notificationEntity.Summary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                ButtonTitle2 = notificationEntity.ButtonTitle2,
                ButtonLink2 = notificationEntity.ButtonLink2,
                CreatedDateTime = notificationEntity.CreatedDate,
                IsRecurrence = notificationEntity.IsRecurrence,
                SentDate = scheduleNotificationEntity.NotificationDate,
                Succeeded = notificationEntity.Succeeded,
                Failed = notificationEntity.Failed,
                Throttled = notificationEntity.Throttled,
                Acknowledged = 0,
                TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Teams),
                RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Rosters),
                AllUsers = notificationEntity.AllUsers,
                SendingStartedDate = notificationEntity.SendingStartedDate,
                Repeats = notificationEntity.Repeats,
                RepeatStartDate = notificationEntity.RepeatStartDate,
                RepeatFor = notificationEntity.RepeatFor,
                RepeatFrequency = notificationEntity.RepeatFrequency,
                WeekSelection = notificationEntity.WeekSelection,
                RepeatEndDate = notificationEntity.RepeatEndDate,
            };

            return this.Ok(result);
        }
    }
}
