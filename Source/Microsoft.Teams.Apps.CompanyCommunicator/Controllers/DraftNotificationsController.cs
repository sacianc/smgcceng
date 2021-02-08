// <copyright file="DraftNotificationsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.NotificationDelivery;
    using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;

    /// <summary>
    /// Controller for the draft notification data.
    /// </summary>
    [Route("api/draftNotifications")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class DraftNotificationsController : ControllerBase
    {
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly TeamDataRepository teamDataRepository;
        private readonly DraftNotificationPreviewService draftNotificationPreviewService;
        private readonly ADGroupsDataRepository adGroupsDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="DraftNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository instance.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        /// <param name="draftNotificationPreviewService">Draft notification preview service.</param>
        /// <param name="adGroupsDataRepository">AD Groups data repository instance.</param>
        public DraftNotificationsController(
            NotificationDataRepository notificationDataRepository,
            TeamDataRepository teamDataRepository,
            DraftNotificationPreviewService draftNotificationPreviewService,
            ADGroupsDataRepository adGroupsDataRepository)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.teamDataRepository = teamDataRepository;
            this.draftNotificationPreviewService = draftNotificationPreviewService;
            this.adGroupsDataRepository = adGroupsDataRepository;
        }

        /// <summary>
        /// Create a new draft notification.
        /// </summary>
        /// <param name="notification">A new Draft Notification to be created.</param>
        /// <returns>The newly created notification's id.</returns>
        [HttpPost]
        public async Task<string> CreateDraftNotificationAsync([FromBody]DraftNotification notification)
        {
            return await this.notificationDataRepository.CreateDraftNotificationAsync(
                notification,
                this.HttpContext.User?.Identity?.Name);
        }

        /// <summary>
        /// Duplicate an existing draft notification.
        /// </summary>
        /// <param name="id">The id of a Draft Notification to be duplicated.</param>
        /// <returns>If the passed in id is invalid, it returns 404 not found error. Otherwise, it returns 200 Ok.</returns>
        [HttpPost("duplicates/{id}")]
        public async Task<IActionResult> DuplicateDraftNotificationAsync(string id)
        {
            var notificationEntity = await this.FindNotificationToDuplicate(id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var createdBy = this.HttpContext.User?.Identity?.Name;

            await this.notificationDataRepository.DuplicateDraftNotificationAsync(notificationEntity, createdBy);

            return this.Ok();
        }

        /// <summary>
        /// Update an existing draft notification.
        /// </summary>
        /// <param name="notification">An existing Draft Notification to be updated.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [HttpPut]
        public async Task UpdateDraftNotificationAsync([FromBody]DraftNotification notification)
        {
            var notificationEntity = new NotificationDataEntity
            {
                PartitionKey = PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                RowKey = notification.Id,
                Id = notification.Id,
                Title = notification.Title,
                ImageLink = notification.ImageLink,
                Summary = notification.Summary,
                Author = notification.Author,
                ButtonTitle = notification.ButtonTitle,
                ButtonLink = notification.ButtonLink,
                ButtonTitle2 = notification.ButtonTitle2,
                ButtonLink2 = notification.ButtonLink2,
                CreatedBy = this.HttpContext.User?.Identity?.Name,
                CreatedDate = DateTime.UtcNow,
                IsDraft = true,
                Teams = notification.Teams,
                Rosters = notification.Rosters,
                ADGroups = notification.ADGroups,
                AllUsers = notification.AllUsers,
                IsScheduled = notification.IsScheduled,
                ScheduleDate = notification.ScheduleDate,
                IsRecurrence = notification.IsRecurrence,
                Repeats = notification.Repeats,
                RepeatFor = Convert.ToInt32(notification.RepeatFor),
                RepeatFrequency = notification.RepeatFrequency,
                WeekSelection = notification.WeekSelection,
                RepeatStartDate = notification.RepeatStartDate,
                RepeatEndDate = notification.RepeatEndDate,
            };

            await this.notificationDataRepository.CreateOrUpdateAsync(notificationEntity);
        }

        /// <summary>
        /// Delete an existing draft notification.
        /// </summary>
        /// <param name="id">The id of the draft notification to be deleted.</param>
        /// <returns>If the passed in Id is invalid, it returns 404 not found error. Otherwise, it returns 200 Ok.</returns>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteDraftNotificationAsync(string id)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            await this.notificationDataRepository.DeleteAsync(notificationEntity);
            return this.Ok();
        }

        /// <summary>
        /// Get draft notifications.
        /// </summary>
        /// <returns>A list of <see cref="DraftNotificationSummary"/> instances.</returns>
        [HttpGet]
        public async Task<ActionResult<IEnumerable<DraftNotificationSummary>>> GetAllDraftNotificationsAsync()
        {
            var notificationEntities = await this.notificationDataRepository.GetAllDraftNotificationsAsync();

            var result = new List<DraftNotificationSummary>();
            foreach (var notificationEntity in notificationEntities)
            {
                var summary = new DraftNotificationSummary
                {
                    Id = notificationEntity.Id,
                    Title = notificationEntity.Title,
                    LastSavedDate = notificationEntity.CreatedDate,
                    IsRecurrence = notificationEntity.IsRecurrence,
                };

                result.Add(summary);
            }

            result.Reverse();

            return result;
        }

        /// <summary>
        /// Get a draft notification by Id.
        /// </summary>
        /// <param name="id">Draft notification Id.</param>
        /// <returns>It returns the draft notification with the passed in id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("{id}")]
        public async Task<ActionResult<DraftNotification>> GetDraftNotificationByIdAsync(string id)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            // AD Groups is empty for old records.
            if (notificationEntity.ADGroups == null)
            {
                notificationEntity.ADGroups = new string[0];
            }

            var result = new DraftNotification
            {
                Id = notificationEntity.Id,
                Title = notificationEntity.Title,
                ImageLink = notificationEntity.ImageLink,
                Summary = notificationEntity.Summary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                CreatedDateTime = notificationEntity.CreatedDate,
                Teams = notificationEntity.Teams,
                Rosters = notificationEntity.Rosters,
                ADGroups = notificationEntity.ADGroups,
                AllUsers = notificationEntity.AllUsers,
                ButtonTitle2 = notificationEntity.ButtonTitle2,
                ButtonLink2 = notificationEntity.ButtonLink2,
                IsScheduled = notificationEntity.IsScheduled,
                ScheduleDate = notificationEntity.ScheduleDate,
                IsRecurrence = notificationEntity.IsRecurrence,
                Repeats = notificationEntity.Repeats,
                RepeatFor = Convert.ToInt32(notificationEntity.RepeatFor),
                RepeatFrequency = notificationEntity.RepeatFrequency,
                WeekSelection = notificationEntity.WeekSelection,
                RepeatStartDate = notificationEntity.RepeatStartDate,
                RepeatEndDate = notificationEntity.RepeatEndDate,
            };

            return this.Ok(result);
        }

        /// <summary>
        /// Get a AD Groups by Search Query.
        /// </summary>
        /// <param name="searchQuery">searchQuery VALUE</param>
        /// <returns>It returns the draft notification with the passed in id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("getADGroups/{searchQuery}")]
        public async Task<ActionResult<ADGroup>> GetADGroupsAsync(string searchQuery)
        {
            if (string.IsNullOrEmpty(searchQuery))
            {
                return this.BadRequest();
            }

            var adGroupEntity = await this.adGroupsDataRepository.GetADGroupsAsync(
                searchQuery);

            if (adGroupEntity == null)
            {
                return this.NotFound();
            }

            return this.Ok(adGroupEntity);
        }

        /// <summary>
        /// Get data for all selected active directory Groups.
        /// </summary>
        /// <param name="adGroups">AD Group Ids.</param>
        /// <returns>A list of active directory groups data.</returns>
        [HttpPost("getADGroupsList")]
        public async Task<IEnumerable<Models.TeamData>> GetADGroupsDataAsync([FromBody]ADGroup[] adGroups)
        {
            List<string> adGroupIds = new List<string>();
            for (int i = 0; i < adGroups.Length; i++)
            {
                adGroupIds.Add(adGroups[i].Id.ToString());
            }

            var entities = await this.adGroupsDataRepository.GetADGroupsList(adGroupIds);
            var result = new List<Models.TeamData>();

            for (int i = 0; i < entities.Count; i++)
            {
                var team = new Models.TeamData
                {
                    TeamId = adGroups[i].Id.ToString(),
                    Name = entities[i].DisplayName.ToString(),
                };
                result.Add(team);
            }

            return result;
        }

        /// <summary>
        /// Get draft notification summary (for consent page) by notification Id.
        /// </summary>
        /// <param name="notificationId">Draft notification Id.</param>
        /// <returns>It returns the draft notification summary (for consent page) with the passed in id.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("consentSummaries/{notificationId}")]
        public async Task<ActionResult<DraftNotificationSummaryForConsent>> GetDraftNotificationSummaryForConsentByIdAsync(string notificationId)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                notificationId);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            List<string> adGroupNames = new List<string>();

            if (notificationEntity.ADGroups.Count() > 0)
            {
                var entities = await this.adGroupsDataRepository.GetADGroupsList(notificationEntity.ADGroups.ToList());

                for (int i = 0; i < entities.Count; i++)
                {
                    adGroupNames.Add(entities[i].DisplayName.ToString());
                }
            }

            var result = new DraftNotificationSummaryForConsent
            {
                NotificationId = notificationId,
                TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Teams),
                RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Rosters),
                ADGroups = adGroupNames,
                AllUsers = notificationEntity.AllUsers,
            };
            return this.Ok(result);
        }

        /// <summary>
        /// Preview draft notification.
        /// </summary>
        /// <param name="draftNotificationPreviewRequest">Draft notification preview request.</param>
        /// <returns>
        /// It returns 400 bad request error if the incoming parameter, draftNotificationPreviewRequest, is invalid.
        /// It returns 404 not found error if the DraftNotificationId or TeamsTeamId (contained in draftNotificationPreviewRequest) is not found in the table storage.
        /// It returns 500 internal error if this method throws an unhandled exception.
        /// It returns 429 too many requests error if the preview request is
        /// by the bot service.
        /// It returns 200 Ok if the method is executed successfully.</returns>
        [HttpPost("previews")]
        public async Task<ActionResult> PreviewDraftNotificationAsync(
            [FromBody] DraftNotificationPreviewRequest draftNotificationPreviewRequest)
        {
            if (draftNotificationPreviewRequest == null
                || string.IsNullOrWhiteSpace(draftNotificationPreviewRequest.DraftNotificationId)
                || string.IsNullOrWhiteSpace(draftNotificationPreviewRequest.TeamsTeamId)
                || string.IsNullOrWhiteSpace(draftNotificationPreviewRequest.TeamsChannelId))
            {
                return this.BadRequest();
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                draftNotificationPreviewRequest.DraftNotificationId);
            if (notificationEntity == null)
            {
                return this.BadRequest($"Notification {draftNotificationPreviewRequest.DraftNotificationId} not found.");
            }

            var teamDataEntity = await this.teamDataRepository.GetAsync(
                PartitionKeyNames.TeamDataTable.TeamDataPartition,
                draftNotificationPreviewRequest.TeamsTeamId);
            if (teamDataEntity == null)
            {
                return this.BadRequest($"Team {draftNotificationPreviewRequest.TeamsTeamId} not found.");
            }

            var result = await this.draftNotificationPreviewService.SendPreview(
                notificationEntity,
                teamDataEntity,
                draftNotificationPreviewRequest.TeamsChannelId);
            return this.StatusCode((int)result);
        }

        private async Task<NotificationDataEntity> FindNotificationToDuplicate(string notificationId)
        {
            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                notificationId);
            if (notificationEntity == null)
            {
                notificationEntity = await this.notificationDataRepository.GetAsync(
                    PartitionKeyNames.NotificationDataTable.SentNotificationsPartition,
                    notificationId);
            }

            if (notificationEntity == null)
            {
                notificationEntity = await this.notificationDataRepository.GetAsync(
                    PartitionKeyNames.NotificationDataTable.ScheduleSentNotificationsPartition,
                    notificationId);
            }

            return notificationEntity;
        }
    }
}