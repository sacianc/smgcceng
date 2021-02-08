// <copyright file="ADGroupsDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Repository of the team data stored in the table storage.
    /// </summary>
    public class ADGroupsDataRepository : BaseRepository<TeamDataEntity>
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private readonly IConfidentialClientApplication app;
        private readonly IConfiguration configuration;
        private readonly string graphQuery = $"https://graph.microsoft.com/v1.0/$batch";
        private readonly string[] scopesVal = { "https://graph.microsoft.com/.default" };

        /// <summary>
        /// Initializes a new instance of the <see cref="ADGroupsDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Represents the application configuration.</param>
        /// <param name="isFromAzureFunction">Flag to show if created from Azure Function.</param>
        public ADGroupsDataRepository(IConfiguration configuration, bool isFromAzureFunction = false)
            : base(
                  configuration,
                  PartitionKeyNames.TeamDataTable.TableName,
                  PartitionKeyNames.TeamDataTable.TeamDataPartition,
                  isFromAzureFunction)
        {
            this.configuration = configuration;
            var microsoftAppId = this.configuration["MicrosoftAppId"];
            var microsoftAppPassword = this.configuration["MicrosoftAppPassword"];
            var aadinstance = this.configuration["AADInstance"];
            var tenantid = this.configuration["TenantId"];

            this.app = ConfidentialClientApplicationBuilder.Create(microsoftAppId)
                .WithClientSecret(microsoftAppPassword)
                .WithAuthority(new Uri(string.Format(CultureInfo.InvariantCulture, aadinstance, tenantid)))
                .Build();
        }

        /// <summary>
        /// Gets team data entities by ID values.
        /// </summary>
        /// <param name="teamIds">Team IDs.</param>
        /// <returns>Team data entities.</returns>
        public async Task<IEnumerable<TeamDataEntity>> GetTeamDataEntitiesByIdsAsync(IEnumerable<string> teamIds)
        {
            var rowKeysFilter = string.Empty;
            foreach (var teamId in teamIds)
            {
                var singleRowKeyFilter = TableQuery.GenerateFilterCondition(
                    nameof(TableEntity.RowKey),
                    QueryComparisons.Equal,
                    teamId);

                if (string.IsNullOrWhiteSpace(rowKeysFilter))
                {
                    rowKeysFilter = singleRowKeyFilter;
                }
                else
                {
                    rowKeysFilter = TableQuery.CombineFilters(rowKeysFilter, TableOperators.Or, singleRowKeyFilter);
                }
            }

            return await this.GetWithFilterAsync(rowKeysFilter);
        }

        /// <summary>
        /// Get team names by Ids.
        /// </summary>
        /// <param name="ids">Team ids.</param>
        /// <returns>Names of the teams matching incoming ids.</returns>
        public async Task<IEnumerable<string>> GetTeamNamesByIdsAsync(IEnumerable<string> ids)
        {
            if (ids == null || ids.Count() == 0)
            {
                return new List<string>();
            }

            var rowKeysFilter = string.Empty;
            foreach (var id in ids)
            {
                var singleRowKeyFilter = TableQuery.GenerateFilterCondition(
                    nameof(TableEntity.RowKey),
                    QueryComparisons.Equal,
                    id);

                if (string.IsNullOrWhiteSpace(rowKeysFilter))
                {
                    rowKeysFilter = singleRowKeyFilter;
                }
                else
                {
                    rowKeysFilter = TableQuery.CombineFilters(rowKeysFilter, TableOperators.Or, singleRowKeyFilter);
                }
            }

            var teamDataEntities = await this.GetWithFilterAsync(rowKeysFilter);

            return teamDataEntities.Select(p => p.Name).OrderBy(p => p);
        }

        /// <summary>
        /// Get all selected AD Group data entities.
        /// </summary>
        /// <param name="adGroupIds">List of AD Group Ids.</param>
        /// <returns>The AD Group data entities.</returns>
        public async Task<List<Models.ADGroup>> GetADGroupsList(IEnumerable<string> adGroupIds)
        {
            string accessToken = this.GetGraphAPIAccessToken<string>();
            List<Models.ADGroup> adGroupMembers = new List<Models.ADGroup>();

            List<Models.ADGroup> responses = await this.GetBatchADGroupsInfoAsync(adGroupIds, accessToken);
            adGroupMembers.AddRange(responses);

            return responses;
        }

        /// <summary>
        /// Gets group members for AD group.
        /// </summary>
        /// <param name="groupIds">Group ids.</param>
        /// <param name="accessToken">accessToken to get Graph API data.</param>
        /// <returns>List of members.</returns>
        public async Task<List<Models.ADGroup>> GetBatchADGroupsInfoAsync(
            IEnumerable<string> groupIds,
            string accessToken)
        {
            var allRequests = new BatchRequestCreator().CreateBatchRequestPayloadForGroupsDetails(groupIds);
            BatchRequestPayload payload = new BatchRequestPayload()
            {
                Requests = allRequests,
            };
            List<Models.ADGroup> responses = await this.CallGraphApiBatchPostToGetGroupsAsync(accessToken, this.graphQuery, JsonConvert.SerializeObject(payload));

            return responses;
        }

        /// <summary>
        /// To Call GraphApi to AD Groups members.
        /// </summary>
        /// <param name="accessToken">Access Token.</param>
        /// <param name="requestUrl">Graph API request URL.</param>
        /// <param name="payload">input JSON.</param>
        /// <returns>A <see cref="T"/> representing the result of operation.</returns>
        public async Task<dynamic> CallGraphApiBatchPostToGetGroupsAsync(string accessToken, string requestUrl, string payload)
        {
            List<Models.ADGroup> groups = new List<Models.ADGroup>();
            HttpClient.DefaultRequestHeaders.Clear();
            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            var request = new StringContent(payload, Encoding.UTF8, "application/json");
            Uri myUri = new Uri(requestUrl);
            HttpResponseMessage response = await HttpClient.PostAsync(myUri, request).ConfigureAwait(true);
            string content = await response.Content.ReadAsStringAsync();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                var responses = this.GetValue<dynamic>(content, "responses");
                foreach (var res in responses)
                {
                    var result = this.GetValue<Models.ADGroup>(JsonConvert.SerializeObject(res), "body");
                    groups.Add(result);
                }

                return groups;
            }

            return content;

            throw new Exception(content);
        }

        /// <summary>
        /// Get AD Groups information.
        /// </summary>
        /// <param name="searchQuery">search Query string.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<List<ADGroup>> GetADGroupsAsync(
            string searchQuery)
        {
            string accessToken = this.GetGraphAPIAccessToken<string>();
            List<ADGroup> adGroups = await this.CallGraphApiToGetGroupsAsync<ADGroup>(accessToken, searchQuery);

            return adGroups;
        }

        /// <summary>
        /// Gets access token and call method to group members for AD groups.
        /// </summary>
        /// <param name="groupIds">Group ids.</param>
        /// <returns>List of members.</returns>
        public async Task<List<ADGroupMembers>> GetADGroupMembersAsync(
            IEnumerable<string> groupIds)
        {
            string accessToken = this.GetGraphAPIAccessToken<string>();
            List<ADGroupMembers> adGroupMembers = new List<ADGroupMembers>();

            List<ADGroupMembers> responses = await this.GetBatchADGroupMembersAsync(groupIds, accessToken);
            adGroupMembers.AddRange(responses);

            return adGroupMembers;
        }

        /// <summary>
        /// Gets group members for AD group.
        /// </summary>
        /// <param name="groupIds">Group ids.</param>
        /// <param name="accessToken">accessToken to get Graph API data.</param>
        /// <returns>List of members.</returns>
        public async Task<List<ADGroupMembers>> GetBatchADGroupMembersAsync(
            IEnumerable<string> groupIds,
            string accessToken)
        {
            List<ADGroupMembers> adGroupMembers = new List<ADGroupMembers>();
            try
            {
                var allRequests = new BatchRequestCreator().CreateBatchRequestPayloadForDetails(groupIds);

                BatchRequestPayload payload = new BatchRequestPayload()
                {
                    Requests = allRequests,
                };
                List<ADGroupMembers> responses = await this.CallGraphApiBatchPostAsync(accessToken, JsonConvert.SerializeObject(payload));
                List<string> nestedGroupIds = new List<string>();

                for (int i = 0; i < responses.Count; i++)
                {
                    if (responses[i].Type == "#microsoft.graph.group")
                    {
                        nestedGroupIds.Add(responses[i].Id);
                    }
                    else
                    {
                        adGroupMembers.Add(responses[i]);
                    }
                }

                if (nestedGroupIds.Count > 0)
                {
                    adGroupMembers.AddRange(await this.GetBatchADGroupMembersAsync(nestedGroupIds, accessToken));
                }
            }
            catch (Exception)
            {
                throw;
            }

            return adGroupMembers;
        }

        /// <summary>
        /// Executes Graph API request.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="accessToken">Access token to authenticate Graph API call.</param>
        /// <param name="query">Graph API query.</param>
        /// <returns>A Response JSON.</returns>
        public async Task<List<ADGroup>> CallGraphApiToGetGroupsAsync<T>(string accessToken, string query)
        {
            string graphQuery = $"https://graph.microsoft.com/v1.0/groups?$top=100&$select=id,displayName,mail&$filter=startswith(displayName,'{query}')";

            List<ADGroup> members = new List<ADGroup>();
            var request = new HttpRequestMessage(HttpMethod.Get, graphQuery);

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            HttpResponseMessage response = await HttpClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();
            if (response.IsSuccessStatusCode)
            {
                if (this.GetValue<List<ADGroup>>(content, "value") != null)
                {
                    members.AddRange(this.GetValue<List<ADGroup>>(content, "value"));
                }
            }

            return members;
        }

        /// <summary>
        /// Executes Graph API request.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="accessToken">Access token.</param>
        /// <param name="groupId">Group id.</param>
        /// <returns>List of users.</returns>
        public async Task<List<ADGroupMembers>> CallGraphApiToGetMembersAsync<T>(string accessToken, string groupId)
        {
            string graphQuery = $"https://graph.microsoft.com/v1.0/groups/{groupId}/members?$select=id,userPrincipalName,department,displayName,givenName,jobTitle,surname";

            List<ADGroupMembers> members = new List<ADGroupMembers>();
            var request = new HttpRequestMessage(HttpMethod.Get, graphQuery);

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            HttpResponseMessage response = await HttpClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();
            if (response.IsSuccessStatusCode)
            {
                if (this.GetValue<List<ADGroupMembers>>(content, "value") != null)
                {
                    members.AddRange(this.GetValue<List<ADGroupMembers>>(content, "value"));
                }
            }

            return members;
        }

        /// <summary>
        /// To Call GraphApi to get AD Group members.
        /// </summary>
        /// <param name="accessToken">Access Token.</param>
        /// <param name="payload">input JSON.</param>
        /// <returns>A <see cref="T"/> representing the result of operation.</returns>
        public async Task<dynamic> CallGraphApiBatchPostAsync(string accessToken, string payload)
        {
            List<ADGroupMembers> users = new List<ADGroupMembers>();
            HttpClient.DefaultRequestHeaders.Clear();
            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            var request = new StringContent(payload, Encoding.UTF8, "application/json");
            Uri myUri = new Uri(this.graphQuery);
            HttpResponseMessage response = await HttpClient.PostAsync(myUri, request).ConfigureAwait(true);
            string content = await response.Content.ReadAsStringAsync();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                var responses = this.GetValue<dynamic>(content, "responses");
                List<string> pagingResultURLs = new List<string>();

                foreach (var res in responses)
                {
                    var userlist = this.GetValue<dynamic>(JsonConvert.SerializeObject(res), "body");
                    var result = this.GetValue<List<ADGroupMembers>>(JsonConvert.SerializeObject(userlist), "value");
                    users.AddRange(result);

                    string query = this.GetValue<string>(JsonConvert.SerializeObject(userlist), "@odata.nextLink");
                    if (query != null)
                    {
                        pagingResultURLs.Add(query);
                    }
                }

                if (pagingResultURLs.Count > 0)
                {
                    var allRequests = new BatchRequestCreator().CreatePagingBatchRequestPayloadForDetails(pagingResultURLs);
                    BatchRequestPayload payloadPagingRequests = new BatchRequestPayload()
                    {
                        Requests = allRequests,
                    };

                    users.AddRange(await this.CallGraphApiBatchPostAsync(accessToken, JsonConvert.SerializeObject(payloadPagingRequests)));
                }

                return users;
            }

            return content;

            throw new Exception(content);
        }

        /// <summary>
        /// To get value from JSON.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="json">Input JSON value.</param>
        /// <param name="jsonPropertyName">Property to retrieve.</param>
        /// <returns>A <see cref="{T}"/> representing the value from JSON.</returns>
        public T GetValue<T>(string json, string jsonPropertyName)
        {
            if (!string.IsNullOrEmpty(json))
            {
                JObject parsedResult = JObject.Parse(json);
                if (parsedResult[jsonPropertyName] != null)
                {
                    return parsedResult[jsonPropertyName].ToObject<T>();
                }
                else
                {
                    return default;
                }
            }
            else
            {
                return default;
            }
        }

        /// <summary>
        /// Get Graph API Access token.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <returns>A Response JSON.</returns>
        public string GetGraphAPIAccessToken<T>()
        {
            AuthenticationResult authResult = this.app.AcquireTokenForClient(this.scopesVal).ExecuteAsync().Result;
            return authResult.AccessToken;
        }

        private class TeamDataEntityComparer : IComparer<TeamDataEntity>
        {
            public int Compare(TeamDataEntity x, TeamDataEntity y)
            {
                return x.Name.CompareTo(y.Name);
            }
        }
    }
}
