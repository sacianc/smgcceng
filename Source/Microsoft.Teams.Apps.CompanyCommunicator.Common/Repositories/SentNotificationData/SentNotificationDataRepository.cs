// <copyright file="SentNotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData
{
    using Microsoft.Extensions.Configuration;
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Repository of the notification data in the table storage.
    /// </summary>
    public class SentNotificationDataRepository : BaseRepository<SentNotificationDataEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SentNotificationDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Represents the application configuration.</param>
        /// <param name="isFromAzureFunction">Flag to show if created from Azure Function.</param>
        public SentNotificationDataRepository(IConfiguration configuration, bool isFromAzureFunction = false)
            : base(
                configuration,
                PartitionKeyNames.SentNotificationDataTable.TableName,
                PartitionKeyNames.SentNotificationDataTable.DefaultPartition,
                isFromAzureFunction)
        {
        }

        /// <summary>
        /// Get acknowledgements.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<IEnumerable<SentNotificationDataEntity>> GetAcknowledgementsAsync()
        {
            var notificationEntity = await this.GetAllAsync(
                PartitionKeyNames.ScheduleNotificationDataTable.ScheduleNotificationsPartition);

            // If Sent notification record deleted from database.
            if (notificationEntity is null)
            {
                return null;
            }

            return notificationEntity;
        }
    }
}