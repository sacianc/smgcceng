// <copyright file="TeamsDataCapture.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;

    /// <summary>
    /// Service to capture teams data.
    /// </summary>
    public class TeamsDataCapture
    {
        private const string PersonalType = "personal";
        private const string ChannelType = "channel";

        private readonly TeamDataRepository teamDataRepository;
        private readonly UserDataRepository userDataRepository;
        private readonly SentNotificationDataRepository sentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamsDataCapture"/> class.
        /// </summary>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        /// <param name="userDataRepository">User data repository instance.</param>
        /// <param name="sentNotificationDataRepository">sentNotificationDataRepository Data Repository instance.</param>
        public TeamsDataCapture(
            TeamDataRepository teamDataRepository,
            UserDataRepository userDataRepository,
            SentNotificationDataRepository sentNotificationDataRepository)
        {
            this.teamDataRepository = teamDataRepository;
            this.userDataRepository = userDataRepository;
            this.sentNotificationDataRepository = sentNotificationDataRepository;
        }

        /// <summary>
        /// Add channel or personal data in Table Storage.
        /// </summary>
        /// <param name="activity">Teams activity instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task OnBotAddedAsync(IConversationUpdateActivity activity)
        {
            switch (activity.Conversation.ConversationType)
            {
                case TeamsDataCapture.ChannelType:
                    await this.teamDataRepository.SaveTeamDataAsync(activity);
                    break;
                case TeamsDataCapture.PersonalType:
                    await this.userDataRepository.SaveUserDataAsync(activity);
                    break;
                default: break;
            }
        }

        /// <summary>
        /// Remove channel or personal data in table storage.
        /// </summary>
        /// <param name="activity">Teams activity instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task OnBotRemovedAsync(IConversationUpdateActivity activity)
        {
            switch (activity.Conversation.ConversationType)
            {
                case TeamsDataCapture.ChannelType:
                    await this.teamDataRepository.RemoveTeamDataAsync(activity);
                    break;
                case TeamsDataCapture.PersonalType:
                    await this.userDataRepository.RemoveUserDataAsync(activity);
                    break;
                default: break;
            }
        }

        /// <summary>
        /// Update team information in the table storage.
        /// </summary>
        /// <param name="activity">Teams activity instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task OnTeamInformationUpdatedAsync(IConversationUpdateActivity activity)
        {
            await this.teamDataRepository.SaveTeamDataAsync(activity);
        }

        /// <summary>
        /// Update Sent notification record information with message acknowledgment flag in the table storage.
        /// </summary>
        /// <param name="notificationId">notificationId to get unique record which is partition key.</param>
        /// <param name="aadObjectId">aadObjectId to get unique record which is row key.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task UpdateRecord(string notificationId, string aadObjectId)
        {
            // Get record from database.
            var sentNotificationDataEntity = await this.sentNotificationDataRepository.GetAsync(
                notificationId,
                aadObjectId);

            // Update with acknowledgment flag.
            sentNotificationDataEntity.MessageAcknowledged = true;

            var operation = TableOperation.InsertOrMerge(sentNotificationDataEntity);
            await this.sentNotificationDataRepository.Table.ExecuteAsync(operation);
        }
    }
}
