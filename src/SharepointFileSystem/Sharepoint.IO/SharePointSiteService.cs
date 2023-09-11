using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint.IO
{
    internal class SharePointSiteService
    {
        //construct a graph client GraphServiceClient
        public SharePointSiteService(GraphServiceClient graphServiceClient)
        {
            GraphServiceClient = graphServiceClient;
        }

        public GraphServiceClient GraphServiceClient { get; }

        ////using the Microsoft.Graph version 5 SDK
        ////use the GraphServiceClient to query the Sites endpoint to find a site by name
        //public Task<Site?> GetSiteAsync(string sitename)
        //{

        //}



    }
}
