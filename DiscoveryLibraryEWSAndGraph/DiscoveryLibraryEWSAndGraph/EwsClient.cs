using System;
using System.Security;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.IdentityModel;

namespace DiscoveryLibraryEWSAndGraph
{
    public class EwsClient : MailboxConnection<MailboxAuthenticationSettings>, MailboxClient
    {
        private ExchangeService Service { get; set; }

        public EwsClient(MailboxAuthenticationSettings mailboxAuthenticationSettings) : base(mailboxAuthenticationSettings)
        {
            Service = new ExchangeService(ExchangeVersion.Exchange2016);
            if(mailboxAuthenticationSettings.Scope == null && mailboxAuthenticationSettings.UserCredential != null)
            {
                Service.Credentials = mailboxAuthenticationSettings.UserCredential;
                if (mailboxAuthenticationSettings.Scope != null)
                {
                    Service.Url = new Uri(mailboxAuthenticationSettings.EndPoint);
                }
                else
                {
                    Service.AutodiscoverUrl(mailboxAuthenticationSettings.UserCredential.UserName);
                }
                
            }
            else
            {
                MSALAppTokenClass mSALAppToken = new MSALAppTokenClass(mailboxAuthenticationSettings);
                Service.Credentials = mSALAppToken;
            }
            Service.Url = new Uri(mailboxAuthenticationSettings.EndPoint);
        }

        public Int64 GetInboxItemCount(string mailboxToAccess)
        {
            Service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxToAccess);
            Service.HttpHeaders.Remove("X-AnchorMailbox");
            Service.HttpHeaders.Add("X-AnchorMailbox", mailboxToAccess);
            var InboxFolder = Folder.Bind(Service, WellKnownFolderName.Inbox);
            return InboxFolder.TotalCount;
        }

    }
}
