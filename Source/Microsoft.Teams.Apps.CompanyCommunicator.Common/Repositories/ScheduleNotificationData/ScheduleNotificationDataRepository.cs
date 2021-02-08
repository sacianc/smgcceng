// <copyright file="ScheduleNotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ScheduleNotificationData
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.NotificationDelivery;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Repository of the schedule notification data in the table storage.
    /// </summary>
    public class ScheduleNotificationDataRepository : BaseRepository<ScheduleNotificationDataEntity>
    {
        private const string EveryWeekday = "Every weekday (Mon-Fri)";
        private const string Daily = "Daily";
        private const string Weekly = "Weekly";
        private const string Monthly = "Monthly";
        private const string Yearly = "Yearly";
        private const string Custom = "Custom";
        private const string Day = "Day";
        private const string Week = "Week";
        private const string Month = "Month";
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly MetadataProvider metadataProvider;
        private readonly SendingNotificationCreator sendingNotificationCreator;
        private readonly ScheduleNotificationDelivery scheduleNotificationDelivery;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScheduleNotificationDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Represents the application configuration.</param>
        /// <param name="notificationDataRepository">Notification DataRepository instance.</param>
        /// <param name="scheduleNotificationDelivery">Notification Delivery instance.</param>
        /// <param name="metadataProvider">Meta data Provider instance.</param>
        /// <param name="sendingNotificationCreator">SendingNotification Creator instance.</param>
        /// <param name="isFromAzureFunction">Flag to show if created from Azure Function.</param>
        public ScheduleNotificationDataRepository(
            IConfiguration configuration,
            NotificationDataRepository notificationDataRepository,
            MetadataProvider metadataProvider = null,
            SendingNotificationCreator sendingNotificationCreator = null,
            ScheduleNotificationDelivery scheduleNotificationDelivery = null,
            bool isFromAzureFunction = false)
            : base(
                configuration,
                PartitionKeyNames.ScheduleNotificationDataTable.TableName,
                PartitionKeyNames.ScheduleNotificationDataTable.ScheduleNotificationsPartition,
                isFromAzureFunction)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.metadataProvider = metadataProvider;
            this.sendingNotificationCreator = sendingNotificationCreator;
            this.scheduleNotificationDelivery = scheduleNotificationDelivery;
        }

        /// <summary>
        /// Create schedule notification record in database.
        /// </summary>
        /// <param name="scheduleNotificationData">The schedule notification record to be created in schedule notification table.</param>
        /// <returns>success or failure flag.</returns>
        public async Task<bool> CreateScheduleNotification(ScheduleNotificationDataEntity scheduleNotificationData)
        {
            if (scheduleNotificationData == null)
            {
                return false;
            }

            await this.CreateOrUpdateAsync(scheduleNotificationData);

            return true;
        }

        /// <summary>
        /// Send a notification to target users.
        /// </summary>
        /// <param name="scheduleNotificationDataEntity">The schedule notification to be sent.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<string> SendNotificationandCreateScheduleAsync(ScheduleNotificationDataEntity scheduleNotificationDataEntity)
        {
            if (scheduleNotificationDataEntity.NotificationId == null)
            {
                return "Notification id empty";
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.ScheduleSentNotificationsPartition,
                scheduleNotificationDataEntity.NotificationId);

            // If Sent notification record deleted from database.
            if (notificationEntity is null)
            {
                return "Notification details does not exist.";
            }

            // Send Notification.
            await this.scheduleNotificationDelivery.SendScheduledNotificationAsync(notificationEntity);

            // If message is recurrence, Generate next schedule.
            if (notificationEntity.IsRecurrence)
            {
                // Create variable with dummy date.
                DateTime notificationDate = DateTime.UtcNow.AddYears(5);
                bool isValidToProceed = true;

                // Calculate next schedule
                switch (notificationEntity.Repeats)
                {
                    case EveryWeekday:
                        notificationDate = scheduleNotificationDataEntity.NotificationDate.AddDays(1);

                        // If next day is Saturday, change to Monday by adding 2
                        if ((int)notificationDate.DayOfWeek == 6)
                        {
                            notificationDate = notificationDate.AddDays(2);
                        }
                        else if ((int)notificationDate.DayOfWeek == 7)
                        {
                            notificationDate = notificationDate.AddDays(1);
                        }

                        break;

                    case Daily:
                        notificationDate = scheduleNotificationDataEntity.NotificationDate.AddDays(1);
                        break;

                    case Weekly:
                        notificationDate = scheduleNotificationDataEntity.NotificationDate.AddDays(7);
                        break;

                    case Monthly:
                        notificationDate = scheduleNotificationDataEntity.NotificationDate.AddMonths(1);
                        break;

                    case Yearly:
                        notificationDate = scheduleNotificationDataEntity.NotificationDate.AddYears(1);
                        break;

                    case Custom:
                        int repeatFor = notificationEntity.RepeatFor;
                        if (notificationEntity.RepeatFrequency == Day)
                        {
                            notificationDate = scheduleNotificationDataEntity.NotificationDate.AddDays(notificationEntity.RepeatFor);
                        }
                        else if (notificationEntity.RepeatFrequency == Month)
                        {
                            notificationDate = scheduleNotificationDataEntity.NotificationDate.AddMonths(notificationEntity.RepeatFor);
                        }
                        else if (notificationEntity.RepeatFrequency == Week)
                        {
                            if (string.IsNullOrEmpty(notificationEntity.WeekSelection))
                            {
                                isValidToProceed = false;
                            }
                            else
                            {
                                string[] weekSelection = notificationEntity.WeekSelection.TrimEnd('/').Split('/');
                                int currentScheduleDayOfWeek = (int)scheduleNotificationDataEntity.NotificationDate.DayOfWeek;
                                for (int i = 0; i < weekSelection.Length; i++)
                                {
                                    if (currentScheduleDayOfWeek == Convert.ToInt32(weekSelection[i]))
                                    {
                                        if (weekSelection.Length > i + 1)
                                        {
                                            notificationDate = scheduleNotificationDataEntity.NotificationDate.AddDays(Convert.ToInt32(weekSelection[i + 1]) - Convert.ToInt32(weekSelection[i]));
                                        }
                                        else
                                        {
                                            notificationDate = scheduleNotificationDataEntity.NotificationDate.AddDays((repeatFor * 7) - Convert.ToInt32(weekSelection[i]) - Convert.ToInt32(weekSelection[0]));
                                        }

                                        break;
                                    }
                                }
                            }
                        }

                        break;
                    default:
                        break;
                }

                if (notificationDate > notificationEntity.RepeatEndDate)
                {
                    isValidToProceed = false;
                }

                if (isValidToProceed)
                {
                    var newScheduleNotificationEntity = new ScheduleNotificationDataEntity
                    {
                        PartitionKey = PartitionKeyNames.ScheduleNotificationDataTable.ScheduleNotificationsPartition,
                        RowKey = scheduleNotificationDataEntity.NotificationId,
                        NotificationId = scheduleNotificationDataEntity.NotificationId,
                        NotificationDate = notificationDate,
                        CreatedDate = DateTime.UtcNow,
                    };

                    // Delete current schedule record.
                    await this.DeleteAsync(scheduleNotificationDataEntity);

                    // Create next schedule record.
                    await this.CreateScheduleNotification(newScheduleNotificationEntity);
                }
                else
                {
                    await this.DeleteAsync(scheduleNotificationDataEntity);
                }
            }
            else
            {
                // Delete current schedule record.
                await this.DeleteAsync(scheduleNotificationDataEntity);
            }

            return string.Empty;
        }

        /// <summary>
        /// Get Scheduled messages.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<IEnumerable<ScheduleNotificationDataEntity>> GetScheduledNotificationsAsync()
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