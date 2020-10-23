// <copyright file="IcebreakerModule.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Icebreaker
{
    using Autofac;
    using Icebreaker.Helpers;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.Extensibility;
    using Microsoft.Azure;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.WebApi;
    using Microsoft.Bot.Connector.Authentication;

    /// <summary>
    /// Autofac Module
    /// </summary>
    public class IcebreakerModule : Module
    {
        /// <inheritdoc/>
        protected override void Load(ContainerBuilder builder)
        {
            base.Load(builder);

            builder.Register(c =>
            {
                return new TelemetryClient(new TelemetryConfiguration(CloudConfigurationManager.GetSetting("APPINSIGHTS_INSTRUMENTATIONKEY")));
            }).SingleInstance();

            builder.Register(c => new MicrosoftAppCredentials(CloudConfigurationManager.GetSetting("MicrosoftAppId"), CloudConfigurationManager.GetSetting("MicrosoftAppPassword")))
                .SingleInstance();

            builder.RegisterType<BotFrameworkHttpAdapter>().As<IBotFrameworkHttpAdapter>()
                .SingleInstance();

            builder.RegisterType<IcebreakerBot>().As<IBot>()
                .SingleInstance();

            builder.RegisterType<IcebreakerBot>()
                .SingleInstance();

            builder.RegisterType<IcebreakerBotDataProvider>()
                .SingleInstance();
        }
    }
}