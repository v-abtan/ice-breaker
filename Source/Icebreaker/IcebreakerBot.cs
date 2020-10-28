//----------------------------------------------------------------------------------------------
// <copyright file="IcebreakerBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------

namespace Icebreaker
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Helpers;
    using Helpers.AdaptiveCards;
    using Icebreaker.Properties;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.WebApi;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Connector.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Implements the core logic for Icebreaker bot
    /// </summary>
    public class IcebreakerBot : TeamsActivityHandler
    {
        private readonly IcebreakerBotDataProvider dataProvider;
        private readonly MicrosoftAppCredentials appCredentials;
        private readonly TelemetryClient telemetryClient;
        private readonly IBotFrameworkHttpAdapter botAdapter;
        private readonly int maxPairUpsPerTeam;
        private readonly string botDisplayName;
        private readonly string botId;
        private readonly bool isTesting;

        /// <summary>
        /// Initializes a new instance of the <see cref="IcebreakerBot"/> class.
        /// </summary>
        /// <param name="dataProvider">The data provider to use</param>
        /// <param name="appCredentials">Microsoft app credentials to use.</param>
        /// <param name="telemetryClient">The telemetry client to use</param>
        /// <param name="botAdapter">Bot adapter.</param>
        public IcebreakerBot(IcebreakerBotDataProvider dataProvider, MicrosoftAppCredentials appCredentials, TelemetryClient telemetryClient, IBotFrameworkHttpAdapter botAdapter)
        {
            this.dataProvider = dataProvider;
            this.appCredentials = appCredentials;
            this.telemetryClient = telemetryClient;
            this.botAdapter = botAdapter;
            this.maxPairUpsPerTeam = Convert.ToInt32(CloudConfigurationManager.GetSetting("MaxPairUpsPerTeam"));
            this.botDisplayName = CloudConfigurationManager.GetSetting("BotDisplayName");
            this.botId = CloudConfigurationManager.GetSetting("MicrosoftAppId");
            this.isTesting = Convert.ToBoolean(CloudConfigurationManager.GetSetting("Testing"));
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onturnasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            try
            {
                this.LogActivityTelemetry(turnContext.Activity);
                await base.OnTurnAsync(turnContext, cancellationToken);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// Conversation update activities are useful when it comes to responding to users being added to or removed from the channel.
        /// For example, a bot could respond to a user being added by greeting the user.
        /// </summary>
        /// <param name="turnContext">A strongly-typed context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// In a derived class, override this method to add logic that applies to all conversation update activities.
        /// </remarks>
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            // conversation-update fires whenever a new 1:1 gets created between us and someone else as well
            // only process the Teams ones.
            var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();
            if (string.IsNullOrEmpty(teamsChannelData?.Team?.Id))
            {
                // conversation-update is for 1:1 chat. Just ignore.
                return;
            }

            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);
        }

        /// <summary>
        /// Provide logic for when members other than the bot
        /// join the conversation, such as your bot's welcome logic.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as
        /// described by the conversation update activity.</param>
        /// <param name="turnContext">A strongly-typed context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// When the <see cref="OnConversationUpdateActivityAsync(ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
        /// method receives a conversation update activity that indicates one or more users other than the bot
        /// are joining the conversation, it calls this method.
        /// </remarks>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            if (membersAdded?.Count() > 0)
            {
                var message = turnContext.Activity;
                string myBotId = message.Recipient.Id;
                string teamId = message.Conversation.Id;
                var teamsChannelData = message.GetChannelData<TeamsChannelData>();

                foreach (var member in membersAdded)
                {
                    if (member.Id == myBotId)
                    {
                        this.telemetryClient.TrackTrace($"Bot installed to team {teamId}");

                        var properties = new Dictionary<string, string>
                                {
                                    { "Scope", message.Conversation?.ConversationType },
                                    { "TeamId", teamId },
                                    { "InstallerId", message.From.Id },
                                };
                        this.telemetryClient.TrackEvent("AppInstalled", properties);

                        // Try to determine the name of the person that installed the app, which is usually the sender of the message (From.Id)
                        // Note that in some cases we cannot resolve it to a team member, because the app was installed to the team programmatically via Graph
                        var personThatAddedBot = (await TeamsInfo.GetMemberAsync(turnContext, message.From.Id, cancellationToken))?.Name;

                        await this.SaveAddedToTeam(message.ServiceUrl, teamId, teamsChannelData.Tenant.Id, personThatAddedBot);
                        await this.WelcomeTeam(turnContext, personThatAddedBot, cancellationToken);
                    }
                    else
                    {
                        this.telemetryClient.TrackTrace($"New member {member.Id} added to team {teamsChannelData.Team.Id}");

                        await this.WelcomeUser(turnContext, member.Id, teamsChannelData.Tenant.Id, teamsChannelData.Team.Id, cancellationToken);
                    }
                }
            }

            await base.OnMembersAddedAsync(membersAdded, turnContext, cancellationToken);
        }

        /// <summary>
        /// Provide logic for when members other than the bot
        /// leave the conversation, such as your bot's good-bye logic.
        /// </summary>
        /// <param name="membersRemoved">A list of all the members removed from the conversation, as
        /// described by the conversation update activity.</param>
        /// <param name="turnContext">A strongly-typed context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// When the <see cref="OnConversationUpdateActivityAsync(ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
        /// method receives a conversation update activity that indicates one or more users other than the bot
        /// are leaving the conversation, it calls this method.
        /// </remarks>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var message = turnContext.Activity;
            string myBotId = message.Recipient.Id;
            string teamId = message.Conversation.Id;
            var teamsChannelData = message.GetChannelData<TeamsChannelData>();
            if (message.MembersRemoved?.Any(x => x.Id == myBotId) == true)
            {
                this.telemetryClient.TrackTrace($"Bot removed from team {teamId}");

                var properties = new Dictionary<string, string>
                {
                    { "Scope", message.Conversation?.ConversationType },
                    { "TeamId", teamId },
                    { "UninstallerId", message.From.Id },
                };
                this.telemetryClient.TrackEvent("AppUninstalled", properties);

                // we were just removed from a team
                await this.SaveRemoveFromTeam(message.ServiceUrl, teamId, teamsChannelData.Tenant.Id);
            }

            await base.OnMembersRemovedAsync(membersRemoved, turnContext, cancellationToken);
        }

        /// <summary>
        /// Provide logic specific to
        /// <see cref="ActivityTypes.Message"/> activities, such as the conversational logic.
        /// Specifically the opt in and out operations.
        /// </summary>
        /// <param name="turnContext">A strongly-typed context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// When the <see cref="OnTurnAsync(ITurnContext, CancellationToken)"/>
        /// method receives a message activity, it calls this method.
        /// </remarks>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            await this.HandleMessageActivityAsync(turnContext, cancellationToken);
            await base.OnMessageActivityAsync(turnContext, cancellationToken);
        }

        /// <summary>
        /// Get a new instance of connector client
        /// </summary>
        /// <param name="serviceUrl">Service url</param>
        /// <returns>connector client instance</returns>
        private IConnectorClient GetConnectorClient(string serviceUrl)
        {
            AppCredentials.TrustServiceUrl(serviceUrl);
            return new ConnectorClient(new Uri(serviceUrl), this.appCredentials);
        }

        /// <summary>
        /// Get TeamsConnectorClient instance from an IConnectorClient.
        /// </summary>
        /// <param name="connectorClient">Generic IConnectorClient instance.</param>
        /// <returns>Returns TeamsConnectorClient to interact with Teams operations.</returns>
        private ITeamsConnectorClient GetTeamsConnectorClient(IConnectorClient connectorClient)
        {
            if (connectorClient is ConnectorClient)
            {
                var connectorClientImpl = (ConnectorClient)connectorClient;
                return new TeamsConnectorClient(connectorClientImpl.BaseUri, connectorClientImpl.Credentials, connectorClientImpl.HttpClient);
            }
            else
            {
                return new TeamsConnectorClient(connectorClient.BaseUri, connectorClient.Credentials);
            }
        }

        /// <summary>
        /// Handle opt in/out operations by updating user preference in data store.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task HandleMessageActivityAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var activity = turnContext.Activity;
                var senderAadId = activity.From.AadObjectId;
                var tenantId = activity.GetChannelData<TeamsChannelData>().Tenant.Id;

                if (string.Equals(activity.Text, "optout", StringComparison.InvariantCultureIgnoreCase))
                {
                    // User opted out
                    this.telemetryClient.TrackTrace($"User {senderAadId} opted out");

                    var properties = new Dictionary<string, string>
                    {
                        { "UserAadId", senderAadId },
                        { "OptInStatus", "false" },
                    };
                    this.telemetryClient.TrackEvent("UserOptInStausSet", properties);

                    await this.OptOutUser(tenantId, senderAadId, activity.ServiceUrl);

                    var optOutReply = activity.CreateReply();
                    optOutReply.Attachments = new List<Attachment>
                    {
                        new HeroCard()
                        {
                            Text = Resources.OptOutConfirmation,
                            Buttons = new List<CardAction>()
                            {
                                new CardAction()
                                {
                                    Title = Resources.ResumePairingsButtonText,
                                    DisplayText = Resources.ResumePairingsButtonText,
                                    Type = ActionTypes.MessageBack,
                                    Text = "optin"
                                }
                            }
                        }.ToAttachment(),
                    };

                    await turnContext.SendActivityAsync(optOutReply, cancellationToken).ConfigureAwait(false);
                }
                else if (string.Equals(activity.Text, "optin", StringComparison.InvariantCultureIgnoreCase))
                {
                    // User opted in
                    this.telemetryClient.TrackTrace($"User {senderAadId} opted in");

                    var properties = new Dictionary<string, string>
                    {
                        { "UserAadId", senderAadId },
                        { "OptInStatus", "true" },
                    };
                    this.telemetryClient.TrackEvent("UserOptInStatusSet", properties);

                    await this.OptInUser(tenantId, senderAadId, activity.ServiceUrl);

                    var optInReply = activity.CreateReply();
                    optInReply.Attachments = new List<Attachment>
                    {
                        new HeroCard()
                        {
                            Text = Resources.OptInConfirmation,
                            Buttons = new List<CardAction>()
                            {
                                new CardAction()
                                {
                                    Title = Resources.PausePairingsButtonText,
                                    DisplayText = Resources.PausePairingsButtonText,
                                    Type = ActionTypes.MessageBack,
                                    Text = "optout"
                                }
                            }
                        }.ToAttachment(),
                    };

                    await turnContext.SendActivityAsync(optInReply, cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    // Unknown input
                    this.telemetryClient.TrackTrace($"Cannot process the following: {activity.Text}");
                    var replyActivity = activity.CreateReply();
                    await this.SendUnrecognizedInputMessageAsync(turnContext, replyActivity, cancellationToken);
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error while handling message activity: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <summary>
        /// Generate pairups and send pairup notifications.
        /// </summary>
        /// <returns>The number of pairups that were made</returns>
        public async Task<int> MakePairsAndNotifyAsync()
        {
            this.telemetryClient.TrackTrace("Making pairups");

            // Recall all the teams where we have been added
            // For each team where bot has been added:
            //     Pull the roster of the team
            //     Remove the members who have opted out of pairups
            //     Match each member with someone else
            //     Save this pair
            // Now notify each pair found in 1:1 and ask them to reach out to the other person
            // When contacting the user in 1:1, give them the button to opt-out
            var installedTeamsCount = 0;
            var pairsNotifiedCount = 0;
            var usersNotifiedCount = 0;

            try
            {
                var teams = await this.dataProvider.GetInstalledTeamsAsync();
                installedTeamsCount = teams.Count;
                this.telemetryClient.TrackTrace($"Generating pairs for {installedTeamsCount} teams");

                foreach (var team in teams)
                {
                    this.telemetryClient.TrackTrace($"Pairing members of team {team.Id}");

                    try
                    {
                        var connectorClient = this.GetConnectorClient(team.ServiceUrl);

                        var teamName = await this.GetTeamNameByIdAsync(connectorClient, team.TeamId);
                        var optedInUsers = await this.GetOptedInUsersAsync(connectorClient, team);

                        foreach (var pair in this.MakePairs(optedInUsers).Take(this.maxPairUpsPerTeam))
                        {
                            usersNotifiedCount += await this.NotifyPairAsync(team, teamName, pair, default(CancellationToken));
                            pairsNotifiedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        this.telemetryClient.TrackTrace($"Error pairing up team members: {ex.Message}", SeverityLevel.Warning);
                        this.telemetryClient.TrackException(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error making pairups: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }

            // Log telemetry about the pairups
            var properties = new Dictionary<string, string>
            {
                { "InstalledTeamsCount", installedTeamsCount.ToString() },
                { "PairsNotifiedCount", pairsNotifiedCount.ToString() },
                { "UsersNotifiedCount", usersNotifiedCount.ToString() },
            };
            this.telemetryClient.TrackEvent("ProcessedPairups", properties);

            this.telemetryClient.TrackTrace($"Made {pairsNotifiedCount} pairups, {usersNotifiedCount} notifications sent");
            return pairsNotifiedCount;
        }

        /// <summary>
        /// Method that will return the information of the installed team
        /// </summary>
        /// <param name="teamId">The team id</param>
        /// <returns>The team that the bot has been installed to</returns>
        public Task<TeamInstallInfo> GetInstalledTeam(string teamId)
        {
            return this.dataProvider.GetInstalledTeamAsync(teamId);
        }

        /// <summary>
        /// Send a welcome message to the user that was just added to a team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="memberAddedId">The id of the added user</param>
        /// <param name="tenantId">The tenant id</param>
        /// <param name="teamId">The id of the team the user was added to</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Tracking task</returns>
        public async Task WelcomeUser(ITurnContext turnContext, string memberAddedId, string tenantId, string teamId, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending welcome message for user {memberAddedId}");

            var installedTeam = await this.GetInstalledTeam(teamId);
            var teamName = turnContext.Activity.TeamsGetTeamInfo().Name;
            ChannelAccount userThatJustJoined = await TeamsInfo.GetMemberAsync(turnContext, memberAddedId, cancellationToken);

            if (userThatJustJoined != null)
            {
                var welcomeMessageCard = WelcomeNewMemberAdaptiveCard.GetCard(teamName, userThatJustJoined.Name, this.botDisplayName, installedTeam.InstallerName);
                await this.NotifyUserAsync(turnContext, welcomeMessageCard, userThatJustJoined, tenantId, cancellationToken);
            }
            else
            {
                this.telemetryClient.TrackTrace($"Member {memberAddedId} was not found in team {teamId}, skipping welcome message.", SeverityLevel.Warning);
            }
        }

        /// <summary>
        /// Sends a welcome message to the General channel of the team that this bot has been installed to
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="botInstaller">The installer of the application</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Tracking task</returns>
        public async Task WelcomeTeam(ITurnContext turnContext, string botInstaller, CancellationToken cancellationToken)
        {
            var teamId = turnContext.Activity.Conversation.Id;
            this.telemetryClient.TrackTrace($"Sending welcome message for team {teamId}");

            var teamName = turnContext.Activity.TeamsGetTeamInfo().Name;
            var welcomeTeamMessageCard = WelcomeTeamAdaptiveCard.GetCard(teamName, this.botDisplayName, botInstaller);
            await this.NotifyTeamAsync(turnContext, welcomeTeamMessageCard, teamId, cancellationToken);
        }

        /// <summary>
        /// Sends a message whenever there is unrecognized input into the bot
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="replyActivity">The activity for replying to a message</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Tracking task</returns>
        public async Task SendUnrecognizedInputMessageAsync(ITurnContext turnContext, Activity replyActivity, CancellationToken cancellationToken)
        {
            var unrecognizedInputAdaptiveCard = UnrecognizedInputAdaptiveCard.GetCard();
            replyActivity.Attachments = new List<Attachment>()
            {
                new Attachment()
                {
                    ContentType = "application/vnd.microsoft.card.adaptive",
                    Content = JsonConvert.DeserializeObject(unrecognizedInputAdaptiveCard)
                }
            };
            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }

        /// <summary>
        /// Save information about the team to which the bot was added.
        /// </summary>
        /// <param name="serviceUrl">The service url</param>
        /// <param name="teamId">The team id</param>
        /// <param name="tenantId">The tenant id</param>
        /// <param name="botInstaller">Person that has added the bot to the team</param>
        /// <returns>Tracking task</returns>
        public Task SaveAddedToTeam(string serviceUrl, string teamId, string tenantId, string botInstaller)
        {
            var teamInstallInfo = new TeamInstallInfo
            {
                ServiceUrl = serviceUrl,
                TeamId = teamId,
                TenantId = tenantId,
                InstallerName = botInstaller
            };
            return this.dataProvider.UpdateTeamInstallStatusAsync(teamInstallInfo, true);
        }

        /// <summary>
        /// Save information about the team from which the bot was removed.
        /// </summary>
        /// <param name="serviceUrl">The service url</param>
        /// <param name="teamId">The team id</param>
        /// <param name="tenantId">The tenant id</param>
        /// <returns>Tracking task</returns>
        public Task SaveRemoveFromTeam(string serviceUrl, string teamId, string tenantId)
        {
            var teamInstallInfo = new TeamInstallInfo
            {
                TeamId = teamId,
                TenantId = tenantId,
            };
            return this.dataProvider.UpdateTeamInstallStatusAsync(teamInstallInfo, false);
        }

        /// <summary>
        /// Opt out the user from further pairups
        /// </summary>
        /// <param name="tenantId">The tenant id</param>
        /// <param name="userId">The user id</param>
        /// <param name="serviceUrl">The service url</param>
        /// <returns>Tracking task</returns>
        public Task OptOutUser(string tenantId, string userId, string serviceUrl)
        {
            return this.dataProvider.SetUserInfoAsync(tenantId, userId, false, serviceUrl);
        }

        /// <summary>
        /// Opt in the user to pairups
        /// </summary>
        /// <param name="tenantId">The tenant id</param>
        /// <param name="userId">The user id</param>
        /// <param name="serviceUrl">The service url</param>
        /// <returns>Tracking task</returns>
        public Task OptInUser(string tenantId, string userId, string serviceUrl)
        {
            return this.dataProvider.SetUserInfoAsync(tenantId, userId, true, serviceUrl);
        }

        /// <summary>
        /// Get the name of a team.
        /// </summary>
        /// <param name="connectorClient">The connector client</param>
        /// <param name="teamId">The team id</param>
        /// <returns>The name of the team</returns>
        private async Task<string> GetTeamNameByIdAsync(IConnectorClient connectorClient, string teamId)
        {
            var teamsConnectorClient = this.GetTeamsConnectorClient(connectorClient);
            var teamDetailsResult = await teamsConnectorClient.Teams.FetchTeamDetailsAsync(teamId);
            return teamDetailsResult.Name;
        }

        /// <summary>
        /// Notify a pairup.
        /// </summary>
        /// <param name="teamModel">DB team model info.</param>
        /// <param name="teamName">MS-Teams team name</param>
        /// <param name="pair">The pairup</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Number of users notified successfully</returns>
        private async Task<int> NotifyPairAsync(TeamInstallInfo teamModel, string teamName, Tuple<ChannelAccount, ChannelAccount> pair, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending pairup notification to {pair.Item1.Id} and {pair.Item2.Id}");

            var teamsPerson1 = JObject.FromObject(pair.Item1).ToObject<TeamsChannelAccount>();
            var teamsPerson2 = JObject.FromObject(pair.Item2).ToObject<TeamsChannelAccount>();

            // Fill in person2's info in the card for person1
            var cardForPerson1 = PairUpNotificationAdaptiveCard.GetCard(teamName, teamsPerson1, teamsPerson2, this.botDisplayName);

            // Fill in person1's info in the card for person2
            var cardForPerson2 = PairUpNotificationAdaptiveCard.GetCard(teamName, teamsPerson2, teamsPerson1, this.botDisplayName);

            // Send notifications and return the number that was successful
            var notifyResults = await Task.WhenAll(
                this.NotifyUserAsync((BotFrameworkHttpAdapter)this.botAdapter, teamModel.ServiceUrl, teamModel.TeamId, cardForPerson1, teamsPerson1, teamModel.TenantId, cancellationToken),
                this.NotifyUserAsync((BotFrameworkHttpAdapter)this.botAdapter, teamModel.ServiceUrl, teamModel.TeamId, cardForPerson2, teamsPerson2, teamModel.TenantId, cancellationToken));
            return notifyResults.Count(wasNotified => wasNotified);
        }

        /// <summary>
        /// Send a card to a user in direct conversation
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cardToSend">The actual welcome card (for the team)</param>
        /// <param name="user">User channel account</param>
        /// <param name="tenantId">Tenant id</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>True/False operation status</returns>
        private async Task<bool> NotifyUserAsync(ITurnContext turnContext, string cardToSend, ChannelAccount user, string tenantId, CancellationToken cancellationToken)
        {
            var botFrameworkAdapter = (BotFrameworkHttpAdapter)turnContext.Adapter;
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            var serviceUrl = turnContext.Activity.ServiceUrl;

            return await this.NotifyUserAsync(botFrameworkAdapter, serviceUrl, teamsChannelId, cardToSend, user, tenantId, cancellationToken);
        }

        /// <summary>
        /// Send a card to a user in direct conversation
        /// </summary>
        /// <param name="botFrameworkAdapter">Bot adapter</param>
        /// <param name="serviceUrl">Service url</param>
        /// <param name="teamsChannelId">Team channel id where the bot is installed</param>
        /// <param name="cardToSend">The actual welcome card (for the team)</param>
        /// <param name="user">User channel account</param>
        /// <param name="tenantId">Tenant id</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>True/False operation status</returns>
        private async Task<bool> NotifyUserAsync(BotFrameworkHttpAdapter botFrameworkAdapter, string serviceUrl, string teamsChannelId, string cardToSend, ChannelAccount user, string tenantId, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending notification to user {user.Id}");

            try
            {
                // construct the activity we want to post
                var welcomeActivity = new Activity()
                {
                    Type = ActivityTypes.Message,
                    Attachments = new List<Attachment>()
                    {
                        new Attachment()
                        {
                            ContentType = "application/vnd.microsoft.card.adaptive",
                            Content = JsonConvert.DeserializeObject(cardToSend),
                        }
                    }
                };

                // conversation parameters
                var conversationParameters = new ConversationParameters
                {
                    Bot = new ChannelAccount { Id = this.botId },
                    Members = new[] { user },
                    ChannelData = new TeamsChannelData
                    {
                        Tenant = new TenantInfo(tenantId),
                    }
                };

                if (!this.isTesting)
                {
                    // shoot the activity over
                    await botFrameworkAdapter.CreateConversationAsync(
                        teamsChannelId,
                        serviceUrl,
                        this.appCredentials,
                        conversationParameters,
                        async (newTurnContext, newCancellationToken) =>
                        {
                            // Get the conversationReference
                            var conversationReference = newTurnContext.Activity.GetConversationReference();

                            // Send the proactive welcome message
                            await botFrameworkAdapter.ContinueConversationAsync(
                                this.appCredentials.MicrosoftAppId,
                                conversationReference,
                                async (conversationTurnContext, conversationCancellationToken) =>
                                {
                                    await conversationTurnContext.SendActivityAsync(welcomeActivity, conversationCancellationToken);
                                },
                                cancellationToken);
                        },
                        cancellationToken).ConfigureAwait(false);
                }

                return true;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error sending notification to user: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
                return false;
            }
        }

        /// <summary>
        /// Method that will send out the message in the General channel of the team
        /// that this bot has been installed to
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cardToSend">The actual welcome card (for the team)</param>
        /// <param name="teamId">The team id</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A tracking task</returns>
        private async Task NotifyTeamAsync(ITurnContext turnContext, string cardToSend, string teamId, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Sending notification to team {teamId}");

            try
            {
                var activity = new Activity()
                {
                    Type = ActivityTypes.Message,
                    Conversation = new ConversationAccount()
                    {
                        Id = teamId
                    },
                    Attachments = new List<Attachment>()
                    {
                        new Attachment()
                        {
                            ContentType = "application/vnd.microsoft.card.adaptive",
                            Content = JsonConvert.DeserializeObject(cardToSend)
                        }
                    }
                };

                var conversationParameters = new ConversationParameters
                {
                    Activity = activity,
                    ChannelData = new TeamsChannelData { Channel = new ChannelInfo(teamId) },
                };

                await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                    null,
                    turnContext.Activity.ServiceUrl,
                    this.appCredentials,
                    conversationParameters,
                    null,
                    cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Error sending notification to team: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <summary>
        /// Get list of opted in users to start matching process
        /// </summary>
        /// <param name="connectorClient">The connector client</param>
        /// <param name="teamInfo">The team that the bot has been installed to</param>
        /// <returns>Opted in users' channels</returns>
        private async Task<List<ChannelAccount>> GetOptedInUsersAsync(IConnectorClient connectorClient, TeamInstallInfo teamInfo)
        {
            // Pull the roster of specified team and then remove everyone who has opted out explicitly
            var members = await connectorClient.Conversations.GetConversationMembersAsync(teamInfo.TeamId);
            this.telemetryClient.TrackTrace($"Found {members.Count} in team {teamInfo.TeamId}");

            var teamMembersIdList = members
                .Where(member => member != null)
                .Select(this.GetChannelUserObjectId)
                .Where(memberObjectId => memberObjectId != null)
                .ToList();
            var dbMembers = (await this.dataProvider.GetUsersInfoAsync(teamMembersIdList))
                .ToDictionary(m => m.Id, m => m.OptedIn);

            return members
                .Where(member => member != null)
                .Where(member =>
                {
                    var memberObjectId = this.GetChannelUserObjectId(member);
                    return !dbMembers.ContainsKey(memberObjectId) || dbMembers[memberObjectId];
                })
                .ToList();
        }

        private string GetChannelUserObjectId(ChannelAccount m)
        {
            return JObject.FromObject(m).ToObject<TeamsChannelAccount>()?.AadObjectId;
        }

        private List<Tuple<ChannelAccount, ChannelAccount>> MakePairs(List<ChannelAccount> users)
        {
            if (users.Count > 1)
            {
                this.telemetryClient.TrackTrace($"Making {users.Count / 2} pairs among {users.Count} users");
            }
            else
            {
                this.telemetryClient.TrackTrace($"Pairs could not be made because there is only 1 user in the team");
            }

            this.Randomize(users);

            var pairs = new List<Tuple<ChannelAccount, ChannelAccount>>();
            for (int i = 0; i < users.Count - 1; i += 2)
            {
                pairs.Add(new Tuple<ChannelAccount, ChannelAccount>(users[i], users[i + 1]));
            }

            return pairs;
        }

        /// <summary>
        /// Randomize list of users
        /// </summary>
        /// <typeparam name="T">Generic item type</typeparam>
        /// <param name="items">List of users to randomize</param>
        private void Randomize<T>(IList<T> items)
        {
            Random rand = new Random(Guid.NewGuid().GetHashCode());

            // For each spot in the array, pick
            // a random item to swap into that spot.
            for (int i = 0; i < items.Count - 1; i++)
            {
                int j = rand.Next(i, items.Count);
                T temp = items[i];
                items[i] = items[j];
                items[j] = temp;
            }
        }

        /// <summary>
        /// Log telemetry about the incoming activity.
        /// </summary>
        /// <param name="activity">The activity</param>
        private void LogActivityTelemetry(Activity activity)
        {
            var fromObjectId = activity.From?.AadObjectId;
            var clientInfoEntity = activity.Entities?.Where(e => e.Type == "clientInfo")?.FirstOrDefault();
            var channelData = activity.GetChannelData<TeamsChannelData>();

            var properties = new Dictionary<string, string>
            {
                { "ActivityId", activity.Id },
                { "ActivityType", activity.Type },
                { "UserAadObjectId", fromObjectId },
                {
                    "ConversationType",
                    string.IsNullOrWhiteSpace(activity.Conversation?.ConversationType) ? "personal" : activity.Conversation.ConversationType
                },
                { "ConversationId", activity.Conversation?.Id },
                { "TeamId", channelData?.Team?.Id },
                { "Locale", clientInfoEntity?.Properties["locale"]?.ToString() },
                { "Platform", clientInfoEntity?.Properties["platform"]?.ToString() }
            };
            this.telemetryClient.TrackEvent("UserActivity", properties);
        }
    }
}