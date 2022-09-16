using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Protocols;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Net;

namespace DiscoveryLibraryEWSAndGraph
{
    public class DiscoveryClient
    {

        private IPublicClientApplication publicClientApplication;
        public string GraphOpenIdDiscovery(string domainName)
        {
            string odicEndpoint = $"https://login.microsoftonline.com/{domainName}/.well-known/openid-configuration";
            var configurationManager = new ConfigurationManager<OpenIdConnectConfiguration>
                (odicEndpoint, new OpenIdConnectConfigurationRetriever(), new HttpDocumentRetriever());
            try
            {
                var odicConfig = configurationManager.GetConfigurationAsync().GetAwaiter().GetResult();
                if(odicConfig.AdditionalData.TryGetValue("msgraph_host",out object msgraph_host_value)){
                    return msgraph_host_value.ToString();
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException is IOException)
                {
                    return null;
                }
                else
                {
                    throw;
                }
            }
            return null;
        }

        public string GetTennantId(string domainName)
        {
            string odicEndpoint = $"https://login.microsoftonline.com/{domainName}/.well-known/openid-configuration";
            var configurationManager = new ConfigurationManager<OpenIdConnectConfiguration>(odicEndpoint, new OpenIdConnectConfigurationRetriever(), new HttpDocumentRetriever());
            try
            {
                var odicConfig = configurationManager.GetConfigurationAsync().GetAwaiter().GetResult();
                return new Uri(odicConfig.Issuer).LocalPath.Replace("/","");
            }
            catch (Exception ex)
            {
                if (ex.InnerException is IOException)
                {
                    return null;
                }
                else
                {
                    throw;
                }
            }
            return null;
        }

        public string AutoDiscoverV2(string emailAddress, string serverEndPoint,bool followRedirect=false)
        {
            var handler = new HttpClientHandler()
            {
                AllowAutoRedirect = followRedirect
            };
            using (HttpClient httpClient = new HttpClient(handler))
            {
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; DicoveryClient/1.0)");
                String autoDiscoverEndpoint = $"https://{serverEndPoint}/autodiscover/autodiscover.json?Email="
                    + Uri.EscapeDataString(emailAddress) + "&Protocol=EWS&RedirectCount=3";
                var adV2Results = httpClient.GetAsync(autoDiscoverEndpoint).GetAwaiter().GetResult();
                if (adV2Results.IsSuccessStatusCode)
                {
                    dynamic jsonResult = JsonConvert.DeserializeObject(httpClient.GetAsync(autoDiscoverEndpoint)
                        .Result.Content.ReadAsStringAsync().Result);
                    if (IsPropertyExist(jsonResult, "Url"))
                    {
                        return jsonResult.Url.ToString();
                    }
                }
            }
            return null;
        }

        public bool CheckForHybridModernAuthentication(string ewsUrl, string tenantId)
        {
            string trustedIssuer = $"00000001-0000-0000-c000-000000000000@{tenantId}";
            using(HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "");
                var authTestResult = client.GetAsync(ewsUrl).GetAwaiter().GetResult();
                foreach(var wwwAuthHeader in authTestResult.Headers.WwwAuthenticate)
                {
                    if(wwwAuthHeader.Scheme == "Bearer" && wwwAuthHeader.Parameter.Contains(trustedIssuer)){
                        return true;
                    }
                }
            }
            return false;
        }
        public AuthenticationResult InstanceAwareInteractiveAuth(string[] scopes,string clientId,string redirectUri)
        {
                    
            PublicClientApplicationBuilder pcaConfig = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority("https://login.microsoftonline.com/common")
                .WithMultiCloudSupport(true)
                .WithRedirectUri(redirectUri);

            publicClientApplication = pcaConfig.Build();
            return InterActiveOAuthAuthentication(scopes);
        }

        public AuthenticationResult AppAuthHybrid(string[] scopes, string clientId,string certificateFileName,SecureString certificatePassword, string tennatId)
        {

            X509Certificate2 x509Certificate2 = new X509Certificate2(certificateFileName,new NetworkCredential("",certificatePassword).Password);
            ConfidentialClientApplicationBuilder ccaConfig = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithTenantId(tennatId)
                .WithCertificate(x509Certificate2);
            IConfidentialClientApplication confidentialClientApplication = ccaConfig.Build();
            return confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();

        }

        public void OpenIdDiscoveryInteractiveAuth(string[] scopes, string clientId, string redirectUri)
        {

            PublicClientApplicationBuilder pcaConfig = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority("https://login.microsoftonline.com/common")
                .WithMultiCloudSupport(true)
                .WithRedirectUri(redirectUri);

            publicClientApplication = pcaConfig.Build();
            var authResult = InterActiveOAuthAuthentication(scopes);
        }

        private AuthenticationResult InterActiveOAuthAuthentication(string[] scopes)
        {
            var accounts = publicClientApplication.GetAccountsAsync().GetAwaiter().GetResult();
            AuthenticationResult result = null;
            try
            {
                result = publicClientApplication.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                                  .ExecuteAsync().GetAwaiter().GetResult();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent.
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                try
                {
                    result = publicClientApplication.AcquireTokenInteractive(scopes)
                                      .ExecuteAsync().GetAwaiter().GetResult();
                }
                catch (MsalException msalex)
                {
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return result;
        }

        public static bool IsPropertyExist(dynamic settings, string name)
        {
            if (settings is Newtonsoft.Json.Linq.JObject)
                return ((Newtonsoft.Json.Linq.JObject)settings).ContainsKey(name);

            return settings.GetType().GetProperty(name) != null;
        }

    }
}
