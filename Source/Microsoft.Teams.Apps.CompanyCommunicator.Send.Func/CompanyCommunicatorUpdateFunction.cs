// <copyright file="CompanyCommunicatorUpdateFunction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Update.Func
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
    public class CompanyCommunicatorUpdateFunction
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
        /// Initializes a new instance of the <see cref="CompanyCommunicatorUpdateFunction"/> class.
        /// </summary>
        public CompanyCommunicatorUpdateFunction()
        {
        }

        /// <summary>
        /// Azure function that finds overdue event occurrences, creates proactive message,
        /// and enqueues the message in the delivery message queue.
        /// </summary>
        /// <param name="scheduleInitializationTimer">The TimerInfo object coming from the Azure Functions system.</param>
        /// <param name="log">The logging service.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [FunctionName("DepartmentUpdateFunction")]
        public async Task RunAsync(
            [TimerTrigger("0 30 9 * * *")]
            TimerInfo scheduleInitializationTimer,
            ILogger log)
        {
            try
            {
                scheduleInitializationTimer = scheduleInitializationTimer ?? throw new ArgumentNullException(nameof(scheduleInitializationTimer));

                log.LogInformation($"Department data update function executed at: {scheduleInitializationTimer.ToString()}.");

                CompanyCommunicatorUpdateFunction.configuration = CompanyCommunicatorUpdateFunction.configuration ??
                    new ConfigurationBuilder()
                        .AddEnvironmentVariables()
                        .Build();

                CompanyCommunicatorUpdateFunction.userDataRepository = CompanyCommunicatorUpdateFunction.userDataRepository
                       ?? new UserDataRepository(CompanyCommunicatorUpdateFunction.configuration, true);

                await CompanyCommunicatorUpdateFunction.userDataRepository.UpdateDepartmentAsync(PartitionKeyNames.UserDataTable.UserDataPartition);

                log.LogInformation($"Department data update function execution completed at: {scheduleInitializationTimer.ToString()}.");
            }
            catch (Exception ex)
            {
                log.LogError($"Department data update failure " + ex.Message, ex);
            }
        }
    }
}