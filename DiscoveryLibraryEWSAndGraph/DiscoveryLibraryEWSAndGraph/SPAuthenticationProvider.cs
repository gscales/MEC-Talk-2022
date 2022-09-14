using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace DiscoveryLibraryEWSAndGraph
{
    public class SPAuthenticationProvider : IAuthenticationProvider
    {

        private readonly IConfidentialClientApplication _app;

        private readonly string _scope;

        public SPAuthenticationProvider(MailboxAuthenticationSettings mailboxAuthenticationSettings, X509Certificate2 certificate)
        {
            _scope = mailboxAuthenticationSettings.Scope;
            _app = ConfidentialClientApplicationBuilder
                .Create(mailboxAuthenticationSettings.ClientId)
                .WithTenantId(mailboxAuthenticationSettings.Tenantid)
                .WithCertificate(certificate)
                .Build();
        }

        public Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var tokenResult = _app.AcquireTokenForClient(new[] { _scope }).ExecuteAsync().Result;
            request.Headers.Add("Authorization", "Bearer " + tokenResult.AccessToken);
            return Task.CompletedTask;
        }
    }
}
