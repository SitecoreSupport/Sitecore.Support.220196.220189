// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ManagerRootsPop3ReceiversCollection.cs" company="Sitecore A/S">
//   Copyright (C) Sitecore A/S. All rights reserved.
// </copyright>
// <summary>
//   The bounced message handler.
// </summary>
// --------------------------------------------------------------------------------------------------------------------
// ReSharper disable once CheckNamespace
using Sitecore.Modules.EmailCampaign;

namespace Sitecore.Support.Modules.EmailCampaign.Core.MessageTransfer
{
    using System;
    using System.Collections.Generic;
    using Diagnostics;
    using Sitecore.EDS.Core.Net.Pop3;
    using Sitecore.EDS.Core.Reporting;
    using Sitecore.EDS.Providers.CustomSmtp.Reporting;
    using ExM.Framework.Diagnostics;

    /// <summary>
    /// The manager roots Pop3 receivers collection.
    /// </summary>
    public class ManagerRootsPop3ReceiversCollection : IPop3ReceiversCollection
    {
        /// <summary>
        /// The logger factory
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// The factory
        /// </summary>
        private readonly Factory factory;

        /// <summary>
        /// The inspector
        /// </summary>
        private readonly Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector;

        /// <summary>
        /// The environment identifier
        /// </summary>
        private readonly IEnvironmentId environmentId;

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagerRootsPop3ReceiversCollection" /> class.
        /// </summary>
        /// <param name="inspector">The inspector.</param>
        /// <param name="environmentId">The environment identifier.</param>
        /// <param name="logger">The logger.</param>
        public ManagerRootsPop3ReceiversCollection([NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId, [NotNull] ILogger logger)
          : this(Factory.Instance, inspector, environmentId, logger)
        {
            Assert.ArgumentNotNull(inspector, "inspector");
            Assert.ArgumentNotNull(environmentId, "environmentId");
            Assert.ArgumentNotNull(logger, "logger");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagerRootsPop3ReceiversCollection" /> class.
        /// </summary>
        /// <param name="factory">The factory.</param>
        /// <param name="inspector">The inspector.</param>
        /// <param name="environmentId">The environment identifier.</param>
        /// <param name="logger">The logger.</param>
        public ManagerRootsPop3ReceiversCollection([NotNull] Factory factory, [NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId, [NotNull] ILogger logger)
        {
            Assert.ArgumentNotNull(factory, "factory");
            Assert.ArgumentNotNull(inspector, "inspector");
            Assert.ArgumentNotNull(environmentId, "environmentId");
            Assert.ArgumentNotNull(logger, "logger");

            this.factory = factory;
            this.inspector = inspector;
            this.environmentId = environmentId;
            this.logger = logger;
        }

        /// <summary>
        /// Gets the receivers list.
        /// </summary>
        /// <returns>
        /// The receivers list <see cref="IEnumerable{IPop3BounceReceiver}" />
        /// </returns>
        [NotNull]
        public IEnumerable<IPop3BounceReceiver> Receivers()
        {
            var receivers = new List<Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver>();
            foreach (var managerRoot in this.factory.GetManagerRoots())
            {
                if (managerRoot.Settings.GatherNotifications)
                {
                    var settings = new Pop3Settings
                    {
                        Password = managerRoot.Settings.POP3Password,
                        Port = managerRoot.Settings.POP3Port,
                        Server = managerRoot.Settings.POP3Server,
                        UseSsl = managerRoot.Settings.POP3SSL,
                        UserName = managerRoot.Settings.POP3UserName,
                        StartTls = !managerRoot.Settings.POP3SSL
                    };

                    try
                    {
                        var receiver = this.CreateReceiver(settings);
                        receivers.Add(receiver);
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex);
                    }
                }
            }

            return receivers;
        }

        private Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver CreateReceiver(Pop3Settings settings)
        {
            Assert.ArgumentNotNullOrEmpty(settings.Server, "settings.Server");
            Assert.ArgumentCondition(settings.Port > 0, "settings.Port", "Missing Port number.");

            return new Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver(settings, this.inspector, this.environmentId, this.logger);
        }
    }
}