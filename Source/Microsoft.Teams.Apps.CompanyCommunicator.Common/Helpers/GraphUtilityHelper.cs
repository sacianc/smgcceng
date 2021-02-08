// <copyright file="GraphUtilityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class will contain Graph SDK read and write operations.
    /// </summary>
    public class GraphUtilityHelper
    {
        private readonly GraphServiceClient graphClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphUtilityHelper"/> class.
        /// </summary>
        /// <param name="accessToken">Token to access MS graph.</param>
        public GraphUtilityHelper(string accessToken)
        {
            this.graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        await Task.Run(() =>
                        {
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue(
                                "Bearer",
                                accessToken);
                        });
                    }));
        }

        /// <summary>
        /// Get User presence details from MS Graph.
        /// </summary>
        /// <param name="usersBatch">List of user entities in batch.</param>
        /// <returns>A collection of user entities with user department information.</returns>
        public async Task<List<UserDataEntity>> GetUserDepartmentAsync(List<UserDataEntity> usersBatch)
        {
            try
            {
                List<UserDataEntity> userDataResults = new List<UserDataEntity>();

                var batchRequestContent = new BatchRequestContent();
                var userIds = usersBatch.Select(user => user.AadId);
                var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("$select", "department,id"),
                    };
                foreach (string userId in userIds)
                {
                    var request = this.graphClient
                        .Users[userId]
                        .Request(queryOptions).GetHttpRequestMessage();
                    request.Method = HttpMethod.Get;

                    batchRequestContent.AddBatchRequestStep(new BatchRequestStep(userId, request));
                }

                var returnedResponse = await this.graphClient.Batch.Request().PostAsync(batchRequestContent);
                var responses = await returnedResponse.GetResponsesAsync();
                foreach (var response in responses)
                {
                    if (response.Value.IsSuccessStatusCode)
                    {
                        var content = await response.Value.Content.ReadAsStringAsync();
                        var responseContent = JsonConvert.DeserializeObject<UserDepartmentResponse>(JObject.Parse(content).ToString());
                        usersBatch.FirstOrDefault(user => user.AadId == responseContent.Id).Department = responseContent.Department;
                    }
                }

                return usersBatch;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
