// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ChilkatPop3BounceReceiver.cs" company="Sitecore A/S">
//   Copyright (C) 2015 by Sitecore A/S
// </copyright>
// <summary>
//   The Chilkat POP3 bounce receiver.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Sitecore.Diagnostics;
using Sitecore.EDS.Core.Diagnostics;
using Sitecore.EDS.Core.Net.Pop3;
using Sitecore.EDS.Core.Reporting;
using Sitecore.ExM.Framework.Diagnostics;
using Sitecore.EDS.Providers.CustomSmtp.Reporting;

namespace Sitecore.Support.EDS.Providers.CustomSmtp.Reporting
{
    /// <summary>
    /// The Chilkat POP3 bounce receiver.
    /// </summary>
    public class ChilkatPop3BounceReceiver : IPop3BounceReceiver
    {
        /// <summary>
        /// The POP3 settings
        /// </summary>
        private readonly Pop3Settings pop3Settings;

        /// <summary>
        /// The inspector.
        /// </summary>
        private readonly Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector;

        /// <summary>
        /// The environment identifier
        /// </summary>
        private readonly IEnvironmentId environmentId;

        /// <summary>
        /// The logger
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChilkatPop3BounceReceiver" /> class.
        /// </summary>
        /// <param name="settings">The POP3 settings.</param>
        /// <param name="inspector">The inspector.</param>
        /// <param name="environmentId">The environment identifier.</param>
        public ChilkatPop3BounceReceiver([NotNull] Pop3Settings settings, [NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId)
          : this(settings, inspector, environmentId, Sitecore.EDS.Core.Diagnostics.LoggerFactory.Instance)
        {
            Assert.ArgumentNotNull(settings, "settings");
            Assert.ArgumentNotNull(inspector, "inspector");
            Assert.ArgumentNotNull(environmentId, "environmentId");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChilkatPop3BounceReceiver" /> class.
        /// </summary>
        /// <param name="settings">The POP3 settings.</param>
        /// <param name="inspector">The inspector.</param>
        /// <param name="environmentId">The environment identifier.</param>
        /// <param name="loggerFactory">The logger factory.</param>
        public ChilkatPop3BounceReceiver([NotNull] Pop3Settings settings, [NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId, [NotNull] ILoggerFactory loggerFactory)
          : this(settings, inspector, environmentId, loggerFactory.Logger)
        {
            Assert.ArgumentNotNull(settings, "settings");
            Assert.ArgumentNotNull(inspector, "inspector");
            Assert.ArgumentNotNull(environmentId, "environmentId");
            Assert.ArgumentNotNull(loggerFactory, "loggerFactory");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChilkatPop3BounceReceiver" /> class.
        /// </summary>
        /// <param name="settings">The POP3 settings.</param>
        /// <param name="inspector">The inspector.</param>
        /// <param name="environmentId">The environment identifier.</param>
        /// <param name="logger">The logger.</param>
        public ChilkatPop3BounceReceiver([NotNull] Pop3Settings settings, [NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId, [NotNull] ILogger logger)
        {
            Assert.ArgumentNotNull(settings, "settings");
            Assert.ArgumentNotNull(inspector, "inspector");
            Assert.ArgumentNotNull(environmentId, "environmentId");
            Assert.ArgumentNotNull(logger, "logger");

            this.pop3Settings = settings;
            this.inspector = inspector;
            this.environmentId = environmentId;
            this.logger = logger;
        }

        /// <summary>
        /// Process the bounced messages.
        /// </summary>
        /// <param name="handleBounces">The bounce handling function.</param>
        /// <returns>
        /// The <see cref="Task" />.
        /// </returns>
        public async Task ProcessMessages(Func<ICollection<Bounce>, Task> handleBounces)
        {
            using (ChilkatPop3Client chilkatPop3Client = (ChilkatPop3Client)this.CreatePop3Client())
            {
                var count = chilkatPop3Client.GetMailboxCount();
                if (count > 0)
                {
                    var messages = chilkatPop3Client.GetMails(false);
                    var bouncedMessages = new List<Bounce>();
                    foreach (var pop3Мail in messages)
                    {
                        var status = this.inspector.InspectEmail(pop3Мail, chilkatPop3Client); // Sitecore.Support.220196.220189
                        if (status != BounceStatus.NotBounce)
                        {
                            var fullМail = chilkatPop3Client.GetMail(pop3Мail.Uidl);
                            if (fullМail == null)
                            {
                                continue;
                            }

                            var environmentIdHeader = fullМail.GetHeader(Sitecore.EDS.Core.Constants.XSitecoreEnvironmentId);
                            if (this.environmentId.IsMatching(environmentIdHeader))
                            {
                                var message = this.MapToBounce(fullМail, status);
                                bouncedMessages.Add(message);
                            }
                        }
                    }

                    if (bouncedMessages.Count > 0)
                    {
                        this.logger.LogInfo(string.Format("ChilkatPop3BounceReceiver processed {0} bounces", bouncedMessages.Count));
                        await handleBounces(bouncedMessages);
                        chilkatPop3Client.DeleteMultipleMails(bouncedMessages.Select(message => message.Id));
                    }
                }
            }
        }

        /// <summary>
        /// Creates the POP3 client.
        /// </summary>
        /// <returns>The POP3 client <see cref="ChilkatPop3Client"/></returns>
        protected virtual IPop3Client CreatePop3Client()
        {
            return new ChilkatPop3Client(this.pop3Settings);
        }

        /// <summary>
        /// Gets the bounce.
        /// </summary>
        /// <param name="pop3Мail">The POP3 mail message.</param>
        /// <param name="bounceType">Bounce Status (Sitecore.Support.220196.220189)</param>
        /// <returns>
        /// Bounce objects.
        /// </returns>
        private Bounce MapToBounce(IPop3Мail pop3Мail, BounceStatus bounceType)
        {
            return new Bounce
            {
                Id = pop3Мail.Uidl,
                MessageId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XBatchId),
                CampaignId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XSitecoreCampaign),
                ContactId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XMessageId),
                // Sitecore.Support.220196.220189: set InstanceId and BounceType
                InstanceId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XBatchId),
                BounceType = bounceType
            };
        }
    }
}
