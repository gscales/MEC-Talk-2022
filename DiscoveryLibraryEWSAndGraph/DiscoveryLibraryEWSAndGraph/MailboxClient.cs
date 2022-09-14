using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace DiscoveryLibraryEWSAndGraph
{
    public interface MailboxClient
    {
        Int64 GetInboxItemCount(string mailboxToAccess);
    }
}
