using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using DiscoveryLibraryEWSAndGraph;

namespace DiscoveryTests
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class ClientTests
    {
        public ClientTests()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void HybridAppAuthAndEWSManaul()
        {
            MailboxAuthenticationSettings mailboxAuthenticationSettings = new MailboxAuthenticationSettings
            {
                EndPoint = "https://exo.mecdemo.msgdevelop.com/ews/exchange.asmx",
                ClientId = "d66f79ab-9457-46ba-b544-abda1ef1e3f4",
                Tenantid = "13af9f3c-b494-4795-bb19-f8364545cd00",
                CertificateFileName = "c:\\temp\\hbc.pfx",
                CertificatePassword = new NetworkCredential("", "xxxx").SecurePassword,
                Scope = "https://exo.mecdemo.msgdevelop.com/.default"
            };
            var ewsClient = new DiscoveryLibraryEWSAndGraph.EwsClient(mailboxAuthenticationSettings);
            var InboxCount = ewsClient.GetInboxItemCount("onpremuser1@mecdemo.msgdevelop.com");
            Assert.AreNotEqual(0,InboxCount);
        }
        [TestMethod]
        public void HybridAppAuthAndEWSAuto()
        {
            MailboxAuthenticationSettings mailboxAuthenticationSettings = new MailboxAuthenticationSettings
            {
                ClientId = "d66f79ab-9457-46ba-b544-abda1ef1e3f4",
                CertificateFileName = "c:\\temp\\hbc.pfx",
                CertificatePassword = new NetworkCredential("", "xxxx").SecurePassword,
            };
            var mailboxToAccess = "onpremuser1@mecdemo.msgdevelop.com";
            var discoveryClient = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            mailboxAuthenticationSettings.Tenantid = discoveryClient.GetTennantId("mecdemo.datarumble.com");
            var url = discoveryClient.AutoDiscoverV2("onpremuser1@mecdemo.msgdevelop.com", "outlook.office365.com", true);
            mailboxAuthenticationSettings.EndPoint = url;
            var aud = new Uri(url).Host;
            mailboxAuthenticationSettings.Scope = $"https://{aud}/.default";
            var ewsClient = new DiscoveryLibraryEWSAndGraph.EwsClient(mailboxAuthenticationSettings);
            var InboxCount = ewsClient.GetInboxItemCount(mailboxToAccess);
            Assert.AreNotEqual(0, InboxCount);
            Console.WriteLine(url);
        }

        [TestMethod]
        public void HybridAppAuthAndGraphAuto()
        {
            MailboxAuthenticationSettings mailboxAuthenticationSettings = new MailboxAuthenticationSettings
            {
                ClientId = "d66f79ab-9457-46ba-b544-abda1ef1e3f4",
                CertificateFileName = "c:\\temp\\hbc.pfx",
                CertificatePassword = new NetworkCredential("", "xxx").SecurePassword,
            };
            var mailboxToAccess = "clouduser1@mecdemo.datarumble.com";
            var discoveryClient = new DiscoveryLibraryEWSAndGraph.DiscoveryClient();
            mailboxAuthenticationSettings.Tenantid = discoveryClient.GetTennantId("mecdemo.datarumble.com");
            var url = discoveryClient.AutoDiscoverV2(mailboxToAccess, "outlook.office365.com", true);
            mailboxAuthenticationSettings.EndPoint = "https://" + discoveryClient.GraphOpenIdDiscovery("mecdemo.datarumble.com");
            var aud = new Uri(mailboxAuthenticationSettings.EndPoint).Host;
            mailboxAuthenticationSettings.Scope = $"https://{aud}/.default";
            var graphClient = new DiscoveryLibraryEWSAndGraph.GraphClient(mailboxAuthenticationSettings);
            var InboxCount = graphClient.GetInboxItemCount(mailboxToAccess);
            Assert.AreNotEqual(0, InboxCount);
            Console.WriteLine(url);
        }
    }
}
