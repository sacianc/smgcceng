// <copyright file="UserDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers;

    /// <summary>
    /// Repository of the user data stored in the table storage.
    /// </summary>
    public class UserDataRepository : BaseRepository<UserDataEntity>
    {
        private const int BatchSplitCount = 20;
        private const int UpdateSplitCount = 100;
        private readonly IConfiguration configuration;
        private readonly string[] scopesVal = { "https://graph.microsoft.com/.default" };

        /// <summary>
        /// Initializes a new instance of the <see cref="UserDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Represents the application configuration.</param>
        /// <param name="isFromAzureFunction">Flag to show if created from Azure Function.</param>
        public UserDataRepository(IConfiguration configuration, bool isFromAzureFunction = false)
            : base(
                configuration,
                PartitionKeyNames.UserDataTable.TableName,
                PartitionKeyNames.UserDataTable.UserDataPartition,
                isFromAzureFunction)
        {
            this.configuration = configuration;
        }

        /// <summary>
        /// Updates department of UserData in Table Storage.
        /// </summary>
        /// <param name="partitionKey">Partition key.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task UpdateDepartmentAsync(string partitionKey)
        {
            try
            {

                if (!string.IsNullOrEmpty(partitionKey))
                {
                    // Get configuration values.
                    var microsoftAppId = this.configuration["MicrosoftAppId"];
                    var microsoftAppPassword = this.configuration["MicrosoftAppPassword"];
                    var aadinstance = this.configuration["AADInstance"];
                    var tenantid = this.configuration["TenantId"];

                    // Generation client application.
                    IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(microsoftAppId)
                        .WithClientSecret(microsoftAppPassword)
                        .WithAuthority(new Uri(string.Format(CultureInfo.InvariantCulture, aadinstance, tenantid)))
                        .Build();

                    // Get all User data records from database. 
                    IEnumerable<UserDataEntity> userDataEntities = await this.GetAllAsync(partitionKey);

                    List<UserDataEntity> usersDataBatchResults = new List<UserDataEntity>();
                    List<UserDataEntity> usersToBeUpdated = new List<UserDataEntity>();

                    // Generate token for Graph API call.
                    string accessToken = this.GetGraphAPIAccessToken<string>(app);
                    GraphUtilityHelper graphClient = new GraphUtilityHelper(accessToken);
                    List<UserDataEntity> userDataEntitiesList = userDataEntities.ToList();
                    if (userDataEntitiesList.Count > 0)
                    {
                        IEnumerable<List<UserDataEntity>> updatedUserSplitList = ListExtensions.SplitList(userDataEntitiesList, BatchSplitCount);
                        foreach (var presenceBatch in updatedUserSplitList)
                        {
                            usersDataBatchResults.AddRange(await graphClient.GetUserDepartmentAsync(presenceBatch));
                        }
                    }

                    foreach (var user in usersDataBatchResults)
                    {
                        if (userDataEntitiesList.FirstOrDefault(x => x.AadId == user.AadId).Department != user.Department)
                        {
                            usersToBeUpdated.Add(user);
                        }
                    }

                    IEnumerable<List<UserDataEntity>> usersSplitList = ListExtensions.SplitList(usersToBeUpdated, UpdateSplitCount);
                    foreach (var updatelist in usersSplitList)
                    {
                        await this.CreateOrUpdateBatchAsync(updatelist);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Get Graph API Access token.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <returns>A Response JSON.</returns>
        public string GetGraphAPIAccessToken<T>(IConfidentialClientApplication app)
        {
            AuthenticationResult authResult = app.AcquireTokenForClient(this.scopesVal).ExecuteAsync().Result;
            return authResult.AccessToken;
        }
    }
}
