using System;
using System.Security;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace DiscoveryLibraryEWSAndGraph
{
    public class MSALAppTokenClass : Microsoft.Exchange.WebServices.Data.Credentials.CustomTokenCredentials
    {
        private string ClientId { get; set; }
        private string TenantId { get; set; }
        private  X509Certificate2 Certificate { get; set; }
        private string Scope { get; set; }       

        public MSALAppTokenClass(MailboxAuthenticationSettings mailboxAuthenticationSettings)
        {
            ClientId = mailboxAuthenticationSettings.ClientId;
            TenantId = mailboxAuthenticationSettings.Tenantid;
            Certificate = new X509Certificate2(mailboxAuthenticationSettings.CertificateFileName, new NetworkCredential("", mailboxAuthenticationSettings.CertificatePassword).Password);
            Scope = mailboxAuthenticationSettings.Scope;
        }
        public IConfidentialClientApplication app { get; set; }
        public override string GetCustomToken()
        {
            if (app == null)
            {
                app = ConfidentialClientApplicationBuilder.Create(ClientId)
                 .WithCertificate(Certificate)
                 .WithTenantId(TenantId)
                 .Build();

            }
            AuthenticationResult result = null;
            try
            {
                result = app.AcquireTokenForClient(new[] { Scope }).ExecuteAsync().Result;
            }
            catch (Exception ex)
            {
                throw;
            }
            return "Bearer " + result.AccessToken;
        }



    }
}
