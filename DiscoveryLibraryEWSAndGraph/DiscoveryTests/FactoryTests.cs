using DiscoveryLibraryEWSAndGraph;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace FactoryTests
{
    [TestClass]
    public class FactoryTests
    {
        [TestMethod]
        public void TestMailClientFactory()
        {
            MailboxAuthenticationSettings mailboxAuthenticationSettings = new MailboxAuthenticationSettings
            {
                ClientId = "d66f79ab-9457-46ba-b544-abda1ef1e3f4",
                Tenantid = "13af9f3c-b494-4795-bb19-f8364545cd00",
                CertificateFileName = "c:\\temp\\hbc.pfx",
                CertificatePassword = new NetworkCredential("", "xxx").SecurePassword,
            };
            string mailboxToAccess = "clouduser1@mecdemo.datarumble.com";
            var mailClient = MailboxClientFactory.GetMailboxClient(mailboxAuthenticationSettings, mailboxToAccess);
            var InboxCount = mailClient.GetInboxItemCount(mailboxToAccess);
            Assert.AreNotEqual(0, InboxCount);
            mailboxToAccess = "onpremuser1@mecdemo.datarumble.com";
            mailClient = MailboxClientFactory.GetMailboxClient(mailboxAuthenticationSettings, mailboxToAccess);
            InboxCount = mailClient.GetInboxItemCount(mailboxToAccess);
            Assert.AreNotEqual(0, InboxCount);
        }
    }    
}
