// <copyright file="BatchRequestCreator.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System.Collections.Generic;

    /// <summary>
    /// creating <see cref="BatchRequestCreator"/> class.
    /// </summary>
    public class BatchRequestCreator
    {
        /// <summary>
        /// Gets or sets this metod to get unique id of each request in batch call.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets this metod to get method of each request in batch call.
        /// </summary>
        public string Method { get; set; }

        /// <summary>
        /// Gets or sets this metod to get url of each request in batch call.
        /// </summary>
        public string URL { get; set; }

        /// <summary>
        /// This method is to create batch request for Graph Api calls.
        /// </summary>
        /// <param name="groupIds">List of strings of groupIds.</param>
        /// <returns>A <see cref="dynamic"/> representing the request payload for batch api call.</returns>
        public dynamic CreateBatchRequestPayloadForDetails(IEnumerable<string> groupIds)
        {
            List<dynamic> request = new List<dynamic>();
            foreach (string id in groupIds)
            {
                BatchRequestCreator batchRequestCreator = new BatchRequestCreator()
                {
                    Id = id,
                    Method = "GET",
                    URL = "groups/" + id + "/members?$top=85&$select=id,userPrincipalName,department,displayName,givenName,jobTitle,surname",
                };
                request.Add(batchRequestCreator);
            }

            return request;
        }

        /// <summary>
        /// This method is to create batch request for Graph Api calls to get paging results.
        /// </summary>
        /// <param name="requesturls">List of strings of request URLs.</param>
        /// <returns>A <see cref="dynamic"/> representing the request payload for batch api call.</returns>
        public dynamic CreatePagingBatchRequestPayloadForDetails(IEnumerable<string> requesturls)
        {
            List<dynamic> request = new List<dynamic>();
            int id = 0;
            foreach (string requestURL in requesturls)
            {
                id++;
                BatchRequestCreator batchRequestCreator = new BatchRequestCreator()
                {
                    Id = id.ToString(),
                    Method = "GET",
                    URL = requestURL.Replace("https://graph.microsoft.com/v1.0", string.Empty),
                };
                request.Add(batchRequestCreator);
            }

            return request;
        }

        /// <summary>
        /// This method is to create batch request for Graph Api calls.
        /// </summary>
        /// <param name="groupIds">List of strings of groupIds.</param>
        /// <returns>A <see cref="dynamic"/> representing the request payload for batch api call.</returns>
        public dynamic CreateBatchRequestPayloadForGroupsDetails(IEnumerable<string> groupIds)
        {
            List<dynamic> request = new List<dynamic>();
            foreach (string id in groupIds)
            {
                BatchRequestCreator batchRequestCreator = new BatchRequestCreator()
                {
                    Id = id,
                    Method = "GET",
                    URL = "groups/" + id + "?$select=id,displayName",
                };
                request.Add(batchRequestCreator);
            }

            return request;
        }
    }
}
