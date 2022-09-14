using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;

namespace DiscoveryLibraryEWSAndGraph
{
    public class MailboxAuthenticationSettings
    {
        public string EndPoint { get; set; }
        public string ClientId { get; set; }
        public string Tenantid { get; set; }
        public string CertificateFileName { get; set; }
        public SecureString CertificatePassword { get; set; }
        public string Scope { get; set; }
    }
}
