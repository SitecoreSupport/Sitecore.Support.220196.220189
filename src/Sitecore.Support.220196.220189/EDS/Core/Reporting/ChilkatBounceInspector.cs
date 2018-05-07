using Chilkat;
using Sitecore.EDS.Core.Net.Pop3;
using Sitecore.EDS.Core.Reporting;
using System;

namespace Sitecore.Support.EDS.Core.Reporting
{

  public class ChilkatBounceInspector : Chilkat.Bounce
  {
    public ChilkatBounceInspector()
    {
      if (!base.UnlockComponent("SITECO.EMX082018_YtCYrNFKpZ5b"))
      {
        throw new Exception("Chilkat Bounce unlock code has expired. Please obtain a new unlock code.");
      }
    }

    protected virtual int InitBounceType(string mimeText)
    {
      base.ExamineMime(mimeText);
      return base.BounceType;
    }

    private BounceStatus InspectMime(string mimeText)
    {
      switch (this.InitBounceType(mimeText))
      {
        case 1:
        case 5:
        case 12:
          return BounceStatus.HardBounce;

        case 2:
        case 4:
        case 11:
          return BounceStatus.SoftBounce;
      }
      return BounceStatus.NotBounce;
    }

    public virtual BounceStatus InspectEmail(IPop3Мail pop3Мail, MailMan mailMan)
    {
      var bounceType = InspectMime(pop3Мail.GetMime);
      if (bounceType == BounceStatus.NotBounce)
      {
        Email email = mailMan.FetchEmail(pop3Мail.Uidl);
        if (email != null )
        {
          if (base.ExamineEmail(email))
          {
            bounceType = MapChilkatBounceToBounce(base.BounceType);
          }
          else
          {
            if (email.IsMultipartReport())
            {
              int statusCode;
              if (int.TryParse(email.GetDeliveryStatusInfo("Status").Replace(".", string.Empty), out statusCode))
              {
                if (statusCode >= 200 && statusCode < 300)
                {
                  bounceType = BounceStatus.NotBounce;
                }
                else if (statusCode >= 400 && statusCode < 500)
                {
                  bounceType = BounceStatus.SoftBounce;
                }
                else if (statusCode >= 500 && statusCode < 600)
                {
                  bounceType = BounceStatus.HardBounce;
                }
                else
                {
                  //unkown or not a bounce code
                }
              }
              else
              {
                //unknown delivery status code
              }
            }
          }
        }
      }
      return bounceType;
    }

    private BounceStatus MapChilkatBounceToBounce(int chilkatBounce)
    {
      switch (chilkatBounce)
      {
        //  Hard bounce
        case 1:
          return BounceStatus.HardBounce;
        //  Soft bounce
        case 2:
        //  General bounce, no email address available.
        case 3:
        //  General bounce
        case 4:
        //  Mail blocked, log the email address
        //  A bounce occured because the sender was blocked.
        case 5:
        //  Transient (recoverable) Failure
        case 7:
        //  Virus Notification
        case 10:
        //  Suspected bounce.
        //  This should be rare.  It indicates that the Bounce
        //  component found strong evidence that this is a bounced
        //  email, but couldn't quite recognize everything it
        //  needed to be 100% sure.
        case 11:
        //  Address Change Notification Message.
        case 13:
          return BounceStatus.SoftBounce;
        //  Auto-reply
        case 6:
        //  Subscribe request
        case 8:
        //  Unsubscribe Request
        case 9:
        //  Challenge/Response - Auto-reply message sent by SPAM software
        //  where only verified email addresses are accepted.
        case 12:
        //  Success DSN indicating that the message was successfully relayed.
        case 14:
          return BounceStatus.NotBounce;
      }

      return BounceStatus.NotBounce;
    }
  }
}
