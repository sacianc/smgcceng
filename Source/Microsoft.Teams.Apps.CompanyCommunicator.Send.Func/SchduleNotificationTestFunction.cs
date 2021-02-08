// <copyright file="SchduleNotificationTestFunction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Functions.SchduleNotification
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.Http;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.NotificationDelivery;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ScheduleNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;

    /// <summary>
    /// Http entry point to the delivery preparation function.
    /// </summary>
    public class SchduleNotificationTestFunction
    {
        private static SendingNotificationDataRepository sendingNotificationDataRepository = null;

        private static NotificationDataRepository notificationDataRepository = null;

        private static GlobalSendingNotificationDataRepository globalSendingNotificationDataRepository = null;

        private static UserDataRepository userDataRepository = null;

        private static TeamDataRepository teamDataRepository = null;

        private static ADGroupsDataRepository adGroupsDataRepository = null;

        private static ScheduleNotificationDelivery scheduleNotificationDelivery = null;

        private static SentNotificationDataRepository sentNotificationDataRepository = null;

        private static MetadataProvider metadataProvider = null;

        private static ScheduleNotificationDataRepository scheduleNotificationDataRepository = null;

        private static SendingNotificationCreator sendingNotificationCreator = null;

        private static IConfiguration configuration = null;

        private static TableRowKeyGenerator tableRowKeyGenerator = null;

        private static AdaptiveCardCreator adaptiveCardCreator = null;

        /// <summary>
        /// Azure function application for testing purpose.
        /// It finds the overdue event occurrences in DB and send them to users.
        /// </summary>
        /// <param name="req">The HTTP request to trigger the function.</param>
        /// <param name="log">The logging service.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName("test-dp")]
        public async Task<IActionResult> RunAsync(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {

          //  scheduleInitializationTimer = scheduleInitializationTimer ?? throw new ArgumentNullException(nameof(scheduleInitializationTimer));

            //log.LogInformation($"Schedule Initialization Timer executed at: {scheduleInitializationTimer.ToString()}.");

            SchduleNotificationTestFunction.configuration = SchduleNotificationTestFunction.configuration ??
                new ConfigurationBuilder()
                    .AddEnvironmentVariables()
                    .Build();

            SchduleNotificationTestFunction.adaptiveCardCreator = SchduleNotificationTestFunction.adaptiveCardCreator
                ?? new AdaptiveCardCreator();

            SchduleNotificationTestFunction.tableRowKeyGenerator = SchduleNotificationTestFunction.tableRowKeyGenerator
                ?? new TableRowKeyGenerator();

            SchduleNotificationTestFunction.adGroupsDataRepository = SchduleNotificationTestFunction.adGroupsDataRepository
                ?? new ADGroupsDataRepository(SchduleNotificationTestFunction.configuration, true);

            SchduleNotificationTestFunction.notificationDataRepository = SchduleNotificationTestFunction.notificationDataRepository
                    ?? new NotificationDataRepository(SchduleNotificationTestFunction.configuration, SchduleNotificationTestFunction.adGroupsDataRepository, tableRowKeyGenerator, true);

            SchduleNotificationTestFunction.userDataRepository = SchduleNotificationTestFunction.userDataRepository
                   ?? new UserDataRepository(SchduleNotificationTestFunction.configuration, true);

            SchduleNotificationTestFunction.teamDataRepository = SchduleNotificationTestFunction.teamDataRepository
                ?? new TeamDataRepository(SchduleNotificationTestFunction.configuration, true);

            SchduleNotificationTestFunction.sendingNotificationDataRepository = SchduleNotificationTestFunction.sendingNotificationDataRepository
                ?? new SendingNotificationDataRepository(SchduleNotificationTestFunction.configuration, true);

            SchduleNotificationTestFunction.globalSendingNotificationDataRepository = SchduleNotificationTestFunction.globalSendingNotificationDataRepository
                ?? new GlobalSendingNotificationDataRepository(SchduleNotificationTestFunction.configuration, true);

            SchduleNotificationTestFunction.sentNotificationDataRepository = SchduleNotificationTestFunction.sentNotificationDataRepository
                ?? new SentNotificationDataRepository(SchduleNotificationTestFunction.configuration, true);

            SchduleNotificationTestFunction.metadataProvider = SchduleNotificationTestFunction.metadataProvider
                ?? new MetadataProvider(SchduleNotificationTestFunction.configuration, userDataRepository, teamDataRepository, notificationDataRepository, adGroupsDataRepository);

            SchduleNotificationTestFunction.sendingNotificationCreator = SchduleNotificationTestFunction.sendingNotificationCreator
                ?? new SendingNotificationCreator(SchduleNotificationTestFunction.configuration, notificationDataRepository, sendingNotificationDataRepository, adaptiveCardCreator);

            SchduleNotificationTestFunction.scheduleNotificationDelivery = SchduleNotificationTestFunction.scheduleNotificationDelivery
                ?? new ScheduleNotificationDelivery(SchduleNotificationTestFunction.configuration, notificationDataRepository, metadataProvider, sendingNotificationCreator, scheduleNotificationDataRepository, teamDataRepository);

            SchduleNotificationTestFunction.scheduleNotificationDataRepository = SchduleNotificationTestFunction.scheduleNotificationDataRepository
                    ?? new ScheduleNotificationDataRepository(SchduleNotificationTestFunction.configuration, notificationDataRepository, metadataProvider, sendingNotificationCreator, scheduleNotificationDelivery, true);

            // Generate filter condition to get pending notifications.
            var rowKeyFilter = TableQuery.GenerateFilterConditionForDate(
                    nameof(ScheduleNotificationDataEntity.NotificationDate),
                    QueryComparisons.LessThanOrEqual,
                    DateTime.UtcNow);

            // Get records which are pending to send notification.
            var pendingNotifications = await SchduleNotificationTestFunction.scheduleNotificationDataRepository.GetWithFilterAsync(
                    rowKeyFilter);

            // Repeat all pending notifications.
            foreach (var pendingNotification in pendingNotifications)
            {
                try
                {
                    log.LogInformation($"Schedule notification triggered for " + pendingNotification.NotificationId);
                    string returnMessage = await SchduleNotificationTestFunction.scheduleNotificationDataRepository.SendNotificationandCreateScheduleAsync(pendingNotification);

                    if (string.IsNullOrEmpty(returnMessage))
                    {
                        log.LogInformation($"Schedule notification success for " + pendingNotification.NotificationId);
                    }
                    else
                    {
                        log.LogInformation($"Schedule notification failed for " + pendingNotification.NotificationId);
                    }
                }
                catch (Exception ex)
                {
                    log.LogError($"Error while sending schedule notification:" + pendingNotification.NotificationId, ex);
                }
            }

            return await Task.FromResult((ActionResult)new OkObjectResult("OK"));
        }
    }
}
