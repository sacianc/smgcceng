// <copyright file="NotificationDelivery.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.NotificationDelivery
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Azure.ServiceBus;
    using Microsoft.Azure.ServiceBus.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ScheduleNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Newtonsoft.Json;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;

    /// <summary>
    /// Notification delivery service.
    /// </summary>
    public class NotificationDelivery
    {
        private readonly IConfiguration configuration;
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly MetadataProvider metadataProvider;
        private readonly SendingNotificationCreator sendingNotificationCreator;
        private readonly ScheduleNotificationDataRepository scheduleNotificationDataRepository;
        private readonly TeamDataRepository teamDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationDelivery"/> class.
        /// </summary>
        /// <param name="configuration">The configuration.</param>
        /// <param name="notificationDataRepository">Notification repository service.</param>
        /// <param name="metadataProvider">Meta data Provider instance.</param>
        /// <param name="sendingNotificationCreator">SendingNotification creator.</param>
        /// <param name="scheduleNotificationDataRepository">Schedule Notification data repository.</param>
        /// <param name="teamDataRepository">Team Data Repository instance.</param>
        public NotificationDelivery(
            IConfiguration configuration,
            NotificationDataRepository notificationDataRepository,
            MetadataProvider metadataProvider,
            SendingNotificationCreator sendingNotificationCreator,
            ScheduleNotificationDataRepository scheduleNotificationDataRepository,
            TeamDataRepository teamDataRepository)
        {
            this.configuration = configuration;
            this.notificationDataRepository = notificationDataRepository;
            this.metadataProvider = metadataProvider;
            this.sendingNotificationCreator = sendingNotificationCreator;
            this.scheduleNotificationDataRepository = scheduleNotificationDataRepository;
            this.teamDataRepository = teamDataRepository;
        }

        /// <summary>
        /// Send a notification to target users.
        /// </summary>
        /// <param name="draftNotificationEntity">The draft notification to be sent.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task SendAsync(NotificationDataEntity draftNotificationEntity)
        {
            if (draftNotificationEntity == null || !draftNotificationEntity.IsDraft)
            {
                return;
            }

            // If message is scheduled or recurrence
            if (draftNotificationEntity.IsScheduled || draftNotificationEntity.IsRecurrence)
            {
                DateTime notificationDate;
                bool isValidToProceed = true;

                // Calculate next schedule
                if (draftNotificationEntity.IsScheduled)
                {
                    notificationDate = draftNotificationEntity.ScheduleDate;
                }
                else
                {
                    DateTime repeatStartDate = draftNotificationEntity.RepeatStartDate;
                    DateTime currentDate = DateTime.UtcNow.AddDays(-1);

                    // If Recurring start date is older than today date, setting to today date
                    if (repeatStartDate < currentDate)
                    {
                        repeatStartDate = currentDate;
                    }

                    notificationDate = repeatStartDate;

                    if (notificationDate > draftNotificationEntity.RepeatEndDate)
                    {
                        isValidToProceed = false;
                    }
                }

                if (isValidToProceed)
                {
                    var newSentNotificationId = await this.notificationDataRepository.MoveDraftToSentPartitionAsync(draftNotificationEntity, true);
                    var scheduleNotificationEntity = new ScheduleNotificationDataEntity
                    {
                        PartitionKey = PartitionKeyNames.ScheduleNotificationDataTable.ScheduleNotificationsPartition,
                        RowKey = newSentNotificationId,
                        NotificationId = newSentNotificationId,
                        NotificationDate = notificationDate,
                        CreatedDate = DateTime.UtcNow,
                    };
                    await this.scheduleNotificationDataRepository.CreateScheduleNotification(scheduleNotificationEntity);
                }
            }
            else // If message is onetime sending
            {
                List<UserDataEntity> deDuplicatedReceiverEntities = new List<UserDataEntity>();

                if (draftNotificationEntity.AllUsers)
                {
                    // Get all users
                    var usersUserDataEntityDictionary = await this.metadataProvider.GetUserDataDictionaryAsync();
                    deDuplicatedReceiverEntities.AddRange(usersUserDataEntityDictionary.Select(kvp => kvp.Value));
                }
                else
                {
                    if (draftNotificationEntity.Rosters.Count() != 0)
                    {
                        var rosterUserDataEntityDictionary = await this.metadataProvider.GetTeamsRostersAsync(draftNotificationEntity.Rosters);

                        deDuplicatedReceiverEntities.AddRange(rosterUserDataEntityDictionary.Select(kvp => kvp.Value));
                    }

                    if (draftNotificationEntity.Teams.Count() != 0)
                    {
                        var teamsReceiverEntities = await this.metadataProvider.GetTeamsReceiverEntities(draftNotificationEntity.Teams);

                        deDuplicatedReceiverEntities.AddRange(teamsReceiverEntities);
                    }

                    if (draftNotificationEntity.ADGroups.Count() != 0)
                    {
                        // Get AD Groups members.
                        var adGroupMemberEntities = await this.metadataProvider.GetADGroupReceiverEntities(draftNotificationEntity.ADGroups);
                        List<UserDataEntity> adGroupMembers = new List<UserDataEntity>();
                        adGroupMembers.AddRange(adGroupMemberEntities);
                        adGroupMembers = adGroupMembers.ToList();

                        // Get all users details from database.
                        var usersUserDataEntityDictionary = await this.metadataProvider.GetUserDataDictionaryAsync();
                        List<UserDataEntity> deAllEntities = new List<UserDataEntity>();
                        deAllEntities.AddRange(usersUserDataEntityDictionary.Select(kvp => kvp.Value));
                        deAllEntities = deAllEntities.ToList();

                        // To get conversation id, mapping all users and ad groups members.
                        for (int i = 0; i < adGroupMembers.Count(); i++)
                        {
                            UserDataEntity userDataEntity = deAllEntities.Find(item => item.AadId == adGroupMembers[i].Id);
                            if (userDataEntity != null && userDataEntity.AadId != null)
                            {
                                deDuplicatedReceiverEntities.Add(userDataEntity);
                            }
                        }

                        deDuplicatedReceiverEntities = deDuplicatedReceiverEntities.Distinct().ToList();
                    }
                }

                var totalMessageCount = deDuplicatedReceiverEntities.Count;
                draftNotificationEntity.TotalMessageCount = totalMessageCount;

                var newSentNotificationId = await this.notificationDataRepository.MoveDraftToSentPartitionAsync(draftNotificationEntity, false);

                // Set in SendingNotification data
                await this.sendingNotificationCreator.CreateAsync(newSentNotificationId, draftNotificationEntity);

                var allServiceBusMessages = deDuplicatedReceiverEntities
                    .Select(userDataEntity =>
                    {
                        var queueMessageContent = new ServiceBusSendQueueMessageContent
                        {
                            NotificationId = newSentNotificationId,
                            UserDataEntity = userDataEntity,
                        };
                        var messageBody = JsonConvert.SerializeObject(queueMessageContent);
                        return new Message(Encoding.UTF8.GetBytes(messageBody));
                    })
                    .ToList();

                // Create batches to send to the service bus
                var serviceBusBatches = new List<List<Message>>();

                var totalNumberMessages = allServiceBusMessages.Count;
                var batchSize = 100;
                var numberOfCompleteBatches = totalNumberMessages / batchSize;
                var numberMessagesInIncompleteBatch = totalNumberMessages % batchSize;

                for (var i = 0; i < numberOfCompleteBatches; i++)
                {
                    var startingIndex = i * batchSize;
                    var batch = allServiceBusMessages.GetRange(startingIndex, batchSize);
                    serviceBusBatches.Add(batch);
                }

                if (numberMessagesInIncompleteBatch != 0)
                {
                    var incompleteBatchStartingIndex = numberOfCompleteBatches * batchSize;
                    var incompleteBatch = allServiceBusMessages.GetRange(
                        incompleteBatchStartingIndex,
                        numberMessagesInIncompleteBatch);
                    serviceBusBatches.Add(incompleteBatch);
                }

                string serviceBusConnectionString = this.configuration["ServiceBusConnection"];
                string queueName = "company-communicator-send";
                var messageSender = new MessageSender(serviceBusConnectionString, queueName);

                // Send batches of messages to the service bus
                foreach (var batch in serviceBusBatches)
                {
                    await messageSender.SendAsync(batch);
                }

                await this.SendTriggerToDataFunction(
                    this.configuration,
                    newSentNotificationId,
                    totalMessageCount);
            }
        }

        private async Task SendTriggerToDataFunction(
            IConfiguration configuration,
            string notificationId,
            int totalMessageCount)
        {
            var queueMessageContent = new ServiceBusDataQueueMessageContent
            {
                NotificationId = notificationId,
                InitialSendDate = DateTime.UtcNow,
                TotalMessageCount = totalMessageCount,
            };
            var messageBody = JsonConvert.SerializeObject(queueMessageContent);
            var serviceBusMessage = new Message(Encoding.UTF8.GetBytes(messageBody));
            serviceBusMessage.ScheduledEnqueueTimeUtc = DateTime.UtcNow + TimeSpan.FromSeconds(30);

            string serviceBusConnectionString = configuration["ServiceBusConnection"];
            string queueName = "company-communicator-data";
            var messageSender = new MessageSender(serviceBusConnectionString, queueName);

            await messageSender.SendAsync(serviceBusMessage);
        }

        private class ServiceBusSendQueueMessageContent
        {
            public string NotificationId { get; set; }

            // This can be a team.id
            public UserDataEntity UserDataEntity { get; set; }
        }

        private class ServiceBusDataQueueMessageContent
        {
            public string NotificationId { get; set; }

            public DateTime InitialSendDate { get; set; }

            public int TotalMessageCount { get; set; }
        }
    }
}
