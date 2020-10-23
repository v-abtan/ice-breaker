//----------------------------------------------------------------------------------------------
// <copyright file="MessagesController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------

using System.Text;

namespace Icebreaker
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading.Tasks;
    using System.Web.Http;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.WebApi;
    using Microsoft.Bot.Connector;
    //using Microsoft.Bot.Connector.Teams.Models;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Properties;

    /// <summary>
    /// Controller for the bot messaging endpoint
    /// </summary>
    public class MessagesController : ApiController
    {
        private readonly IBotFrameworkHttpAdapter adapter;
        private readonly IBot bot;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessagesController"/> class.
        /// </summary>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="bot">The Icebreaker bot instance</param>
        public MessagesController(IBotFrameworkHttpAdapter adapter, IBot bot)
        {
            this.adapter = adapter;
            this.bot = bot;
        }

        public async Task<HttpResponseMessage> Post()
        {
            var response = new HttpResponseMessage();

            // Delegate the processing of the HTTP POST to the adapter.
            // The adapter will invoke the bot.
            await this.adapter.ProcessAsync(this.Request, response, this.bot);
            return response;
        }

        /*
        /// <summary>
        /// POST: api/messages
        /// Receive a message from a user and reply to it
        /// </summary>
        /// <param name="activity">The incoming activity</param>
        /// <returns>Task that resolves to the HTTP response message</returns>
        public async Task<HttpResponseMessage> Post([FromBody] Activity activity)
        {
            this.LogActivityTelemetry(activity);

            using (var connectorClient = new ConnectorClient(new Uri(activity.ServiceUrl)))
            {
                if (activity.Type == ActivityTypes.Message)
                {
                    await this.HandleMessageActivity(connectorClient, activity);
                }
                else
                {
                    await this.HandleSystemActivity(connectorClient, activity);
                }
            }

            return this.Request.CreateResponse(HttpStatusCode.OK);
        }

        
        private async Task HandleSystemActivity(ConnectorClient connectorClient, Activity message)
        {
            this.telemetryClient.TrackTrace("Processing system message");

            try
            {
                var teamsChannelData = message.GetChannelData<TeamsChannelData>();
                var tenantId = teamsChannelData.Tenant.Id;

                if (message.Type == ActivityTypes.ConversationUpdate)
                {
        
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error while handling system activity: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }
        */

    }
}