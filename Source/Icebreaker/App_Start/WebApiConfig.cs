// <copyright file="WebApiConfig.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Icebreaker
{
    using System.Reflection;
    using System.Web.Http;
    using System.Web.Http.Dependencies;
    using Autofac;
    using Autofac.Integration.WebApi;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Connector.Authentication;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Serialization;

    /// <summary>
    /// Web API configuration
    /// </summary>
    public static class WebApiConfig
    {
        /// <summary>
        /// Configures API settings
        /// </summary>
        /// <param name="config">The <see cref="HttpConfiguration"/> to configure</param>
        public static void Register(HttpConfiguration config)
        {
            // Json settings
            config.Formatters.JsonFormatter.SerializerSettings.NullValueHandling = NullValueHandling.Ignore;
            config.Formatters.JsonFormatter.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
            config.Formatters.JsonFormatter.SerializerSettings.Formatting = Formatting.Indented;
            JsonConvert.DefaultSettings = () => new JsonSerializerSettings()
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                Formatting = Formatting.Indented,
                NullValueHandling = NullValueHandling.Ignore,
            };

            // Web API configuration and services
            config.DependencyResolver = GetResolver();

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional });
        }

        private static IDependencyResolver GetResolver()
        {
            var builder = new ContainerBuilder();
            builder.RegisterApiControllers(Assembly.GetExecutingAssembly());
            builder.RegisterWebApiFilterProvider(GlobalConfiguration.Configuration);

            // The ConfigurationCredentialProvider will retrieve the MicrosoftAppId and
            // MicrosoftAppPassword from Web.config
            builder.RegisterType<ConfigurationCredentialProvider>().As<ICredentialProvider>().SingleInstance();

            // Create the Bot Framework Adapter with error handling enabled.
            //builder.RegisterType<AdapterWithErrorHandler>().As<IBotFrameworkHttpAdapter>().SingleInstance();

            // The Memory Storage used here is for local bot debugging only. When the bot
            // is restarted, everything stored in memory will be gone.
            IStorage dataStore = new MemoryStorage();

            // Create Conversation State object.
            // The Conversation State object is where we persist anything at the conversation-scope.
            var conversationState = new ConversationState(dataStore);
            builder.RegisterInstance(conversationState).As<ConversationState>().SingleInstance();

            // Register the main dialog, which is injected into the DialogBot class
            //builder.RegisterType<RootDialog>().SingleInstance();

            // Register the DialogBot with RootDialog as the IBot interface
            //builder.RegisterType<DialogBot<RootDialog>>().As<IBot>();

            builder.RegisterModule(new IcebreakerModule());

            var container = builder.Build();
            var resolver = new AutofacWebApiDependencyResolver(container);
            return resolver;
        }
    }
}
