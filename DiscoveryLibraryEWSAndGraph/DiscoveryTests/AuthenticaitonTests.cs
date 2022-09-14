using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Net;
using System.Collections.Generic;
using System.Text;

namespace DiscoveryTests
{
    [TestClass]
    public class AuthenticaitonTests
    {
        [TestMethod]
        public void InstanceAwareInterActiveAuthentication()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var authResult = client.InstanceAwareInteractiveAuth(new[] { "Mail.Read" }, "f6d34fc3-bfc5-41c4-94e0-bdb99fb67cb2", "https://login.microsoftonline.com/common/oauth2/nativeclient");
            Assert.IsNotNull(authResult.AccessToken);
        }

        [TestMethod]
        public void HybridAppAuth()
        {
            var client = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var authResult = client.AppAuthHybrid(new[] { "https://exotest.mecdemo.msgdevelop.com/.default" }, "d66f79ab-9457-46ba-b544-abda1ef1e3f4"
                , "c:\\temp\\hbc.pfx",new NetworkCredential("","xxx").SecurePassword , "13af9f3c-b494-4795-bb19-f8364545cd00");
            Assert.IsNotNull(authResult.AccessToken);
        }
        [TestMethod]
        public void CheckForHybridModernAuth()
        {
            var discoveryClient = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            var hasHMA = discoveryClient.CheckForHybridModernAuthentication("onpremuser1@mecdemo.msgdevelop.com", "13af9f3c-b494-4795-bb19-f8364545cd00");
            Assert.IsTrue(hasHMA);

        }
    }
}
