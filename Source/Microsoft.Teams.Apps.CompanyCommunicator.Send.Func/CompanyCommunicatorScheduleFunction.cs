// <copyright file="CompanyCommunicatorScheduleFunction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Schedule.Func
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Azure.WebJobs;
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
    /// Azure function application triggered by timer hourly.
    /// It finds the overdue event occurrences in DB and send them to users.
    /// </summary>
    public class CompanyCommunicatorScheduleFunction
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
        /// Initializes a new instance of the <see cref="CompanyCommunicatorScheduleFunction"/> class.
        /// </summary>
        public CompanyCommunicatorScheduleFunction()
        {
        }

        /// <summary>
        /// Azure function that finds overdue event occurrences, creates proactive message,
        /// and enqueues the message in the delivery message queue.
        /// </summary>
        /// <param name="scheduleInitializationTimer">The TimerInfo object coming from the Azure Functions system.</param>
        /// <param name="log">The logging service.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [FunctionName("ScheduleFunction")]
        public async Task RunAsync(
            [TimerTrigger("0 */30 * * * *")]
            TimerInfo scheduleInitializationTimer,
            ILogger log)
        {
            scheduleInitializationTimer = scheduleInitializationTimer ?? throw new ArgumentNullException(nameof(scheduleInitializationTimer));

            log.LogInformation($"Schedule Initialization Timer executed at: {scheduleInitializationTimer.ToString()}.");

            CompanyCommunicatorScheduleFunction.configuration = CompanyCommunicatorScheduleFunction.configuration ??
                new ConfigurationBuilder()
                    .AddEnvironmentVariables()
                    .Build();

            CompanyCommunicatorScheduleFunction.adaptiveCardCreator = CompanyCommunicatorScheduleFunction.adaptiveCardCreator
                ?? new AdaptiveCardCreator();

            CompanyCommunicatorScheduleFunction.tableRowKeyGenerator = CompanyCommunicatorScheduleFunction.tableRowKeyGenerator
                ?? new TableRowKeyGenerator();

            CompanyCommunicatorScheduleFunction.adGroupsDataRepository = CompanyCommunicatorScheduleFunction.adGroupsDataRepository
                ?? new ADGroupsDataRepository(CompanyCommunicatorScheduleFunction.configuration, true);

            CompanyCommunicatorScheduleFunction.notificationDataRepository = CompanyCommunicatorScheduleFunction.notificationDataRepository
                    ?? new NotificationDataRepository(CompanyCommunicatorScheduleFunction.configuration, CompanyCommunicatorScheduleFunction.adGroupsDataRepository, tableRowKeyGenerator, true);

            CompanyCommunicatorScheduleFunction.userDataRepository = CompanyCommunicatorScheduleFunction.userDataRepository
                   ?? new UserDataRepository(CompanyCommunicatorScheduleFunction.configuration, true);

            CompanyCommunicatorScheduleFunction.teamDataRepository = CompanyCommunicatorScheduleFunction.teamDataRepository
                ?? new TeamDataRepository(CompanyCommunicatorScheduleFunction.configuration, true);

            CompanyCommunicatorScheduleFunction.sendingNotificationDataRepository = CompanyCommunicatorScheduleFunction.sendingNotificationDataRepository
                ?? new SendingNotificationDataRepository(CompanyCommunicatorScheduleFunction.configuration, true);

            CompanyCommunicatorScheduleFunction.globalSendingNotificationDataRepository = CompanyCommunicatorScheduleFunction.globalSendingNotificationDataRepository
                ?? new GlobalSendingNotificationDataRepository(CompanyCommunicatorScheduleFunction.configuration, true);

            CompanyCommunicatorScheduleFunction.sentNotificationDataRepository = CompanyCommunicatorScheduleFunction.sentNotificationDataRepository
                ?? new SentNotificationDataRepository(CompanyCommunicatorScheduleFunction.configuration, true);

            CompanyCommunicatorScheduleFunction.metadataProvider = CompanyCommunicatorScheduleFunction.metadataProvider
                ?? new MetadataProvider(CompanyCommunicatorScheduleFunction.configuration, userDataRepository, teamDataRepository, notificationDataRepository, adGroupsDataRepository);

            CompanyCommunicatorScheduleFunction.sendingNotificationCreator = CompanyCommunicatorScheduleFunction.sendingNotificationCreator
                ?? new SendingNotificationCreator(CompanyCommunicatorScheduleFunction.configuration, notificationDataRepository, sendingNotificationDataRepository, adaptiveCardCreator);

            CompanyCommunicatorScheduleFunction.scheduleNotificationDelivery = CompanyCommunicatorScheduleFunction.scheduleNotificationDelivery
                ?? new ScheduleNotificationDelivery(CompanyCommunicatorScheduleFunction.configuration, notificationDataRepository, metadataProvider, sendingNotificationCreator, scheduleNotificationDataRepository, teamDataRepository);

            CompanyCommunicatorScheduleFunction.scheduleNotificationDataRepository = CompanyCommunicatorScheduleFunction.scheduleNotificationDataRepository
                    ?? new ScheduleNotificationDataRepository(CompanyCommunicatorScheduleFunction.configuration, notificationDataRepository, metadataProvider, sendingNotificationCreator, scheduleNotificationDelivery, true);

            // Generate filter condition to get pending notifications.
            var rowKeyFilter = TableQuery.GenerateFilterConditionForDate(
                    nameof(ScheduleNotificationDataEntity.NotificationDate),
                    QueryComparisons.LessThanOrEqual,
                    DateTime.UtcNow);

            // Get records which are pending to send notification.
            var pendingNotifications = await CompanyCommunicatorScheduleFunction.scheduleNotificationDataRepository.GetWithFilterAsync(
                    rowKeyFilter);

            // Repeat all pending notifications.
            foreach (var pendingNotification in pendingNotifications)
            {
                try
                {
                    log.LogInformation($"Schedule notification triggered for " + pendingNotification.NotificationId);
                    string returnMessage = await CompanyCommunicatorScheduleFunction.scheduleNotificationDataRepository.SendNotificationandCreateScheduleAsync(pendingNotification);

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
        }
    }
}