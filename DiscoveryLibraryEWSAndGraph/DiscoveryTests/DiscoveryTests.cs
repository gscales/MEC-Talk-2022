using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Net;

namespace DiscoveryTests
{
    [TestClass]
    public class DiscoveryTests
    {
        [TestMethod]
        public void OpenTestGlobalDomain()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var graphhost = client.GraphOpenIdDiscovery("contoso.com");
            Assert.AreEqual(graphhost, "graph.microsoft.com");
        }
        [TestMethod]
        public void OpenTestNationDomain()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var graphhost = client.GraphOpenIdDiscovery("contosochina.com");
            Assert.AreEqual(graphhost, "microsoftgraph.chinacloudapi.cn");
        }
        [TestMethod]
        public void OpenIdValidDomain()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var graphhost = client.GraphOpenIdDiscovery("datarumble.com");
            Assert.AreEqual(graphhost, "graph.microsoft.com");
        }
        [TestMethod]
        public void OpenIdValidNoOffice365()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var graphhost = client.GraphOpenIdDiscovery("contso.msgdevelop.com");
            Assert.IsNull(graphhost);
        }
        [TestMethod]
        public void Office365AutoDiscover()
        {
            ServicePointManager
    .ServerCertificateValidationCallback +=
    (sender, cert, chain, sslPolicyErrors) => true;
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var url = client.AutoDiscoverV2("clouduser1@mecdemo.msgdevelop.com", "outlook.office365.com",true);
            Assert.AreEqual(url, "https://outlook.office365.com/EWS/Exchange.asmx");
        }

        [TestMethod]
        public void Office365AutoDiscoverInvalidEmail()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var url = client.AutoDiscoverV2("blahblah@nodomain.com", "outlook.office365.com");
            Assert.IsNull(url);
        }


    }
}
