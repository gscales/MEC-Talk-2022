using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using Microsoft.Graph;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IdentityModel;

namespace DiscoveryLibraryEWSAndGraph
{
    public class GraphClient : MailboxConnection<MailboxAuthenticationSettings>, MailboxClient
    {
        private GraphServiceClient GraphServiceClient { get;set;}

        public GraphClient(MailboxAuthenticationSettings mailboxAuthenticationSettings) : base(mailboxAuthenticationSettings)
        {
            var certificate = new X509Certificate2(mailboxAuthenticationSettings.CertificateFileName, new NetworkCredential("", mailboxAuthenticationSettings.CertificatePassword).Password);
            GraphServiceClient = new GraphServiceClient(new SPAuthenticationProvider(mailboxAuthenticationSettings, certificate));
        }

        public Int64 GetInboxItemCount(string mailboxToAccess)
        {        
            var inboxFolder = GraphServiceClient.Users[mailboxToAccess].MailFolders["inbox"].Request().GetAsync().GetAwaiter().GetResult();
            return (long)inboxFolder.TotalItemCount;
        }
    }
}
