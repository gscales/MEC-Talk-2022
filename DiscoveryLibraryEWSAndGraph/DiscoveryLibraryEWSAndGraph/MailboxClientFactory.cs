using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiscoveryLibraryEWSAndGraph
{
    public class MailboxClientFactory
    {
        public static MailboxClient GetMailboxClient(MailboxAuthenticationSettings mailboxAuthenticationSettings,string mailboxToAccess)
        {
            DiscoveryClient discoveryClient = new DiscoveryClient();
            var graphEndPoint = discoveryClient.GraphOpenIdDiscovery(mailboxToAccess.Split('@')[1]);
            if (graphEndPoint != null)
            {
                var ewsURL = discoveryClient.AutoDiscoverV2(mailboxToAccess, Utils.GraphToO365Endpoint(graphEndPoint), true);
                if (new Uri(ewsURL).Host == Utils.GraphToO365Endpoint(graphEndPoint))
                {
                    mailboxAuthenticationSettings.EndPoint = $"https://{graphEndPoint}";
                    mailboxAuthenticationSettings.Scope = $"https://{graphEndPoint}/.default";
                    return new GraphClient(mailboxAuthenticationSettings);
                }
                else
                {
                    ewsURL = discoveryClient.AutoDiscoverV2(mailboxToAccess, Utils.GraphToO365Endpoint(graphEndPoint), true);
                    if (discoveryClient.CheckForHybridModernAuthentication(ewsURL, mailboxAuthenticationSettings.Tenantid))
                    {
                        mailboxAuthenticationSettings.EndPoint = ewsURL;
                        string aud = new Uri(ewsURL).Host;
                        mailboxAuthenticationSettings.Scope = $"https://{aud}/.default";
                        return new EwsClient(mailboxAuthenticationSettings);
                    }
                    else
                    {
                        mailboxAuthenticationSettings.Scope = null;
                        mailboxAuthenticationSettings.EndPoint = ewsURL;
                        return new EwsClient(mailboxAuthenticationSettings);
                    }
                }
            }
            else
            {
                mailboxAuthenticationSettings.Scope = null;
                return new EwsClient(mailboxAuthenticationSettings);
            }
        }
    }
}
