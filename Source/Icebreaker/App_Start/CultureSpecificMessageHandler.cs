// <copyright file="CultureSpecificMessageHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Icebreaker
{
    using System.Globalization;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure;

    /// <summary>
    /// Message handler to set culture specific settings
    /// </summary>
    public class CultureSpecificMessageHandler : DelegatingHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var cultureName = CloudConfigurationManager.GetSetting("DefaultCulture");
            Thread.CurrentThread.CurrentCulture = new CultureInfo(cultureName);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(cultureName);
            return base.SendAsync(request, cancellationToken);
        }
    }
}