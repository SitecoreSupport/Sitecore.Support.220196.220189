// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ChilkatBounceInspector.cs" company="Sitecore A/S">
//   Copyright (C) 2015 by Sitecore A/S
// </copyright>
// <summary>
//   The chilkat bounce inspector.
//   Verifies the message mime is a bounce and determines the type of the bounce.
// </summary>
// --------------------------------------------------------------------------------------------------------------------
using Chilkat;
using Sitecore.EDS.Core.Net.Pop3;
using Sitecore.EDS.Core.Reporting;

namespace Sitecore.Support.EDS.Core.Reporting
{
    using System;

    /// <summary>
    /// The chilkat bounce inspector.
    /// Verifies the message mime is a bounce and determines the type of the bounce.
    /// </summary>
    public class ChilkatBounceInspector : Chilkat.Bounce, IBounceInspector
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChilkatBounceInspector" /> class.
        /// </summary>
        /// <exception cref="System.Exception">Chilkat Bounce unlock code has expired. Please obtain a new unlock code.</exception>
        public ChilkatBounceInspector()
        {
            if (!this.UnlockComponent(Sitecore.EDS.Core.Constants.MailmanKey))
            {
                throw new Exception("Chilkat Bounce unlock code has expired. Please obtain a new unlock code.");
            }
        }

        /// <summary>
        /// The examine mail.
        /// </summary>
        /// <param name="mimeText">The email mime.</param>
        /// <returns>
        /// The <see cref="BounceStatus" />.
        /// </returns>
        public BounceStatus InspectMime(string mimeText)
        {
            var bounceType = this.InitBounceType(mimeText);

            switch (bounceType)
            {
                // Hard Bounce. The email could not be delivered and BounceAddress contains the failed email address. 
                case 1:
                // Mail Block. A bounce occured because the sender was blocked. 
                case 5:
                // Challenge/Response - Auto-reply message sent by SPAM software where only verified email addresses are accepted.
                case 12:
                    return BounceStatus.HardBounce;

                // Soft Bounce. A temporary condition exists causing the email delivery to fail. The BounceAddress property contains the failed email address. 
                case 2:
                // General Bounced Mail, cannot determine if it is hard or soft, but an email address is available. 
                case 4:
                // Suspected Bounce, but no other information is available 
                case 11:
                    return BounceStatus.SoftBounce;
            }

            return BounceStatus.NotBounce;
        }

        /// <summary>
        /// Initializes the type of the bounce.
        /// </summary>
        /// <param name="mimeText">The MIME text.</param>
        /// <returns>A number representing the type of bounce that was recognized.</returns>
        protected virtual int InitBounceType(string mimeText)
        {
            this.ExamineMime(mimeText);
            return this.BounceType;
        }

        public virtual BounceStatus InspectEmail(IPop3Мail pop3Мail, MailMan mailMan)
        {
            BounceStatus bounceStatus = this.InspectMime(pop3Мail.GetMime);
            if (bounceStatus == BounceStatus.NotBounce)
            {
                Email email = mailMan.FetchEmail(pop3Мail.Uidl);
                if (email != null)
                {
                    int num;
                    if (base.ExamineEmail(email))
                    {
                        bounceStatus = this.MapChilkatBounceToBounce(base.BounceType);
                    }
                    else if (email.IsMultipartReport() && int.TryParse(email.GetDeliveryStatusInfo("Status").Replace(".", string.Empty), out num))
                    {
                        if (num >= 200 && num < 300)
                        {
                            bounceStatus = BounceStatus.NotBounce;
                        }
                        else if (num >= 400 && num < 500)
                        {
                            bounceStatus = BounceStatus.SoftBounce;
                        }
                        else if (num >= 500 && num < 600)
                        {
                            bounceStatus = BounceStatus.HardBounce;
                        }
                    }
                }
            }
            return bounceStatus;
        }

        private BounceStatus MapChilkatBounceToBounce(int chilkatBounce)
        {
            switch (chilkatBounce)
            {
                case 1:
                    return BounceStatus.HardBounce;
                case 2:
                case 3:
                case 4:
                case 5:
                case 7:
                case 10:
                case 11:
                case 13:
                    return BounceStatus.SoftBounce;
                case 6:
                case 8:
                case 9:
                case 12:
                case 14:
                    return BounceStatus.NotBounce;
                default:
                    return BounceStatus.NotBounce;
            }
        }
    }
}
