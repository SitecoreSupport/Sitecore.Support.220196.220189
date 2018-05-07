using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Sitecore.Diagnostics;
using Sitecore.EDS.Core.Diagnostics;
using Sitecore.EDS.Core.Net.Pop3;
using Sitecore.EDS.Core.Reporting;
using Sitecore.EDS.Providers.CustomSmtp.Reporting;
using Sitecore.ExM.Framework.Diagnostics;

namespace Sitecore.Support.EDS.Providers.CustomSmtp.Reporting
{
  public class ChilkatPop3BounceReceiver : IPop3BounceReceiver
  {
    private readonly Pop3Settings pop3Settings;

    private readonly Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector;

    private readonly IEnvironmentId environmentId;

    private readonly ILogger logger;

    public ChilkatPop3BounceReceiver([NotNull] Pop3Settings settings, [NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId)
      : this(settings, inspector, environmentId, Sitecore.EDS.Core.Diagnostics.LoggerFactory.Instance)
    {
      Assert.ArgumentNotNull(settings, "settings");
      Assert.ArgumentNotNull(inspector, "inspector");
      Assert.ArgumentNotNull(environmentId, "environmentId");
    }

    public ChilkatPop3BounceReceiver([NotNull] Pop3Settings settings, [NotNull] Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, [NotNull] IEnvironmentId environmentId, [NotNull] ILoggerFactory loggerFactory)
      : this(settings, inspector, environmentId, loggerFactory.Logger)
    {
      Assert.ArgumentNotNull(settings, "settings");
      Assert.ArgumentNotNull(inspector, "inspector");
      Assert.ArgumentNotNull(environmentId, "environmentId");
      Assert.ArgumentNotNull(loggerFactory, "loggerFactory");
    }

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

    public async Task ProcessMessages(Func<ICollection<Bounce>, Task> handleBounces)
    {
      using (var pop3Client = this.CreatePop3Client())
      {
        var count = pop3Client.GetMailboxCount();
        if (count > 0)
        {
          var messages = pop3Client.GetMails(false);
          List<Bounce> bouncedMessages = new List<Bounce>();
          foreach (var pop3Мail in messages)
          {
            var status = this.inspector.InspectEmail(pop3Мail, pop3Client);
            if (status != BounceStatus.NotBounce)
            {
              var fullМail = pop3Client.GetMail(pop3Мail.Uidl);
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
            pop3Client.DeleteMultipleMails(bouncedMessages.Select(message => message.Id));
          }
        }
      }
    }

    protected virtual ChilkatPop3Client CreatePop3Client()
    {
      return new ChilkatPop3Client(this.pop3Settings);
    }

    private Bounce MapToBounce(IPop3Мail pop3Мail, BounceStatus bounceType)
    {
      return new Bounce
      {
        Id = pop3Мail.Uidl,
        MessageId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XBatchId),
        CampaignId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XSitecoreCampaign),
        ContactId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XMessageId),
        InstanceId = pop3Мail.GetHeader(Sitecore.EDS.Core.Constants.XBatchId),
        BounceType = bounceType
      };
    }
  }
}
