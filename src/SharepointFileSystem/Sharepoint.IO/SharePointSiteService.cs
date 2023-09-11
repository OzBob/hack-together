using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint.IO
{
    public interface ISharePointSiteService
    {
        Task<Site?> GetSiteByNameAsync(string sitename);
        Task<Site?> GetSiteBySiteIdOrFullPathAsync(string siteIdOrFullPath);
        Task<Site?> GetSiteSubSiteAsync(string siteId, string subSiteName);
    }

    public class SharePointSiteService : ISharePointSiteService
    {
        //construct a graph client GraphServiceClient
        public SharePointSiteService(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        public GraphServiceClient _graphServiceClient { get; }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a site by name
        public async Task<Site?> GetSiteByNameAsync(string sitename)
        {
            var site = await _graphServiceClient
                 .Sites[$"name eq {sitename}"]
                    .GetAsync();

            if (site == null)
            {
                Trace.WriteLine($"No Site({sitename}");
            }
            return site;
        }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a site by name
        public async Task<Site?> GetSiteBySiteIdOrFullPathAsync(string siteIdOrFullPath)
        {
            var site = await _graphServiceClient
                 .Sites[$"{siteIdOrFullPath}"]
                    .GetAsync();

            if (site == null)
            {
                Trace.WriteLine($"No Site({siteIdOrFullPath}");
            }
            return site;
        }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a subsite by name
        public async Task<Site?> GetSiteSubSiteAsync(string siteId, string subSiteName)
        {
            var site = await _graphServiceClient
                .Sites[$"{siteId}"].GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "sites" };
                });

            if (site == null)
            {
                Trace.WriteLine($"No main Site({siteId}");
                return null;
            }

            if (site.Sites == null || site.Sites.Count == 0)
            {
                Trace.WriteLine($"No Subsites on Site({siteId})");
                return null;
            }

            var subsite = site.Sites.Where(s => s.Name == subSiteName).FirstOrDefault();
            if (subsite == null || subsite == default)
            {
                Trace.WriteLine($"No Subsite({subSiteName})");
            }
            return subsite;
        }
    }
}
