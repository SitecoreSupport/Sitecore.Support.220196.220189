namespace Sitecore.Support.Modules.EmailCampaign.Core.MessageTransfer
{
  using Microsoft.Extensions.DependencyInjection;
  using Sitecore.DependencyInjection;
  using Sitecore.Diagnostics;
  using Sitecore.EDS.Core.Net.Pop3;
  using Sitecore.EDS.Core.Reporting;
  using Sitecore.EDS.Providers.CustomSmtp.Reporting;
  using Sitecore.ExM.Framework.Diagnostics;
  using Sitecore.Modules.EmailCampaign;
  using Sitecore.Modules.EmailCampaign.Services;
  using System;
  using System.Collections.Generic;

  public class ManagerRootsPop3ReceiversCollection : IPop3ReceiversCollection
  {
    private readonly IManagerRootService _managerRootService;
    private readonly IEnvironmentId environmentId;
    private readonly Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector;
    private readonly ILogger logger;

    public ManagerRootsPop3ReceiversCollection(Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, IEnvironmentId environmentId, ILogger logger) : this(ServiceProviderServiceExtensions.GetService<IManagerRootService>(ServiceLocator.ServiceProvider), inspector, environmentId, logger)
    {
      Assert.ArgumentNotNull(inspector, "inspector");
      Assert.ArgumentNotNull(environmentId, "environmentId");
      Assert.ArgumentNotNull(logger, "logger");
    }

    public ManagerRootsPop3ReceiversCollection(IManagerRootService managerRootService, Sitecore.Support.EDS.Core.Reporting.ChilkatBounceInspector inspector, IEnvironmentId environmentId, ILogger logger)
    {
      Assert.ArgumentNotNull(managerRootService, "managerRootService");
      Assert.ArgumentNotNull(inspector, "inspector");
      Assert.ArgumentNotNull(environmentId, "environmentId");
      Assert.ArgumentNotNull(logger, "logger");
      this._managerRootService = managerRootService;
      this.inspector = inspector;
      this.environmentId = environmentId;
      this.logger = logger;
    }

    private Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver CreateReceiver(Pop3Settings settings)
    {
      Assert.ArgumentNotNullOrEmpty(settings.Server, "settings.Server");
      Assert.ArgumentCondition(settings.Port > 0, "settings.Port", "Missing Port number.");
      return new Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver(settings, this.inspector, this.environmentId, this.logger);
    }

    public IEnumerable<IPop3BounceReceiver> Receivers()
    {
      List<Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver> list = new List<Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver>();
      foreach (ManagerRoot root in this._managerRootService.GetManagerRoots())
      {
        if (root.Settings.GatherNotifications)
        {
          Pop3Settings settings = new Pop3Settings
          {
            Password = root.Settings.POP3Password,
            Port = root.Settings.POP3Port,
            Server = root.Settings.POP3Server,
            UseSsl = root.Settings.POP3SSL,
            UserName = root.Settings.POP3UserName,
            StartTls = !root.Settings.POP3SSL
          };
          try
          {
            Sitecore.Support.EDS.Providers.CustomSmtp.Reporting.ChilkatPop3BounceReceiver item = this.CreateReceiver(settings);
            list.Add(item);
          }
          catch (Exception exception)
          {
            this.logger.LogError(exception);
          }
        }
      }
      return (IEnumerable<IPop3BounceReceiver>)list;
    }
  }
}
