using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiscoveryLibraryEWSAndGraph
{
    internal class Utils
    {
        public static string GraphToO365Endpoint(string graphEndPoint)
        {
            switch (graphEndPoint)
            {
                case "microsoftgraph.chinacloudapi.cn":
                    return "partner.outlook.cn";
                case "graph.microsoft.de":
                    return "outlook.office.de";
                case "dod-graph.microsoft.com":
                    return "outlook-dod.office365.us";
                case "graph.microsoft.us":
                    return "outlook.office365.us";
                default: 
                return "outlook.office365.com";
            }
        }
    }
}
