// <copyright file="ScheduleNotificationDelivery.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.NotificationDelivery
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Azure.ServiceBus;
    using Microsoft.Azure.ServiceBus.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ScheduleNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Newtonsoft.Json;

    /// <summary>
    /// Notification delivery service.
    /// </summary>
    public class ScheduleNotificationDelivery
    {
        private readonly IConfiguration configuration;
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly MetadataProvider metadataProvider;
        private readonly SendingNotificationCreator sendingNotificationCreator;
        private readonly ScheduleNotificationDataRepository scheduleNotificationDataRepository;
        private readonly TeamDataRepository teamDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScheduleNotificationDelivery"/> class.
        /// </summary>
        /// <param name="configuration">The configuration.</param>
        /// <param name="notificationDataRepository">Notification repository service.</param>
        /// <param name="metadataProvider">Meta data Provider instance.</param>
        /// <param name="sendingNotificationCreator">SendingNotification creator.</param>
        /// <param name="scheduleNotificationDataRepository">Schedule Notification data repository.</param>
        /// <param name="teamDataRepository">TeamData Repository instance.</param>
        public ScheduleNotificationDelivery(
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
        /// Send a scheduled notification to target users.
        /// </summary>
        /// <param name="notificationEntity">The notification to be sent.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task SendScheduledNotificationAsync(NotificationDataEntity notificationEntity)
        {
            List<UserDataEntity> deDuplicatedReceiverEntities = new List<UserDataEntity>();

            if (notificationEntity.AllUsers)
            {
                // Get all users
                var usersUserDataEntityDictionary = await this.metadataProvider.GetUserDataDictionaryAsync();
                deDuplicatedReceiverEntities.AddRange(usersUserDataEntityDictionary.Select(kvp => kvp.Value));
            }
            else
            {
                if (notificationEntity.Rosters.Count() != 0)
                {
                    var rosterUserDataEntityDictionary = await this.metadataProvider.GetTeamsRostersAsync(notificationEntity.Rosters);

                    deDuplicatedReceiverEntities.AddRange(rosterUserDataEntityDictionary.Select(kvp => kvp.Value));
                }

                if (notificationEntity.Teams.Count() != 0)
                {
                    var teamsReceiverEntities = await this.metadataProvider.GetTeamsReceiverEntities(notificationEntity.Teams);

                    deDuplicatedReceiverEntities.AddRange(teamsReceiverEntities);
                }

                if (notificationEntity.ADGroups.Count() != 0)
                {
                    // Get AD Groups members.
                    var adGroupMemberEntities = await this.metadataProvider.GetADGroupReceiverEntities(notificationEntity.ADGroups);
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
            notificationEntity.TotalMessageCount = totalMessageCount;

            // Creates record in Sent notifications.
            var newSentNotificationId = await this.notificationDataRepository.CopyToSentPartitionAsync(notificationEntity);

            // Set in SendingNotification data
            await this.sendingNotificationCreator.CreateAsync(newSentNotificationId, notificationEntity);

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
