using Azure;
using Microsoft.Graph.Models;
using System.Linq;

namespace Sharepoint.IO.Model
{
    public class SpSite
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? WebUrl { get; set; }
        //public string? Root { get; set; }
        public string? SiteCollectionRoot { get; set; }
        public string? SiteCollectionHostname { get; set; }
        public string? SiteUrl { get; set; }
        public SpSite()
        {
            BaseDriveFolder = new SpFolder();
        }
        public SpSite(Site site)
        {
            Id = site.Id;
            Name = site.Name;
            WebUrl = site.WebUrl;
            //SiteCollectionRoot = site.SiteCollection.Root.BackingStore.;
            SiteCollectionHostname = site.SiteCollection?.Hostname;
            SiteUrl = site.SharepointIds?.SiteUrl;
            BaseDriveFolder = new SpFolder();
        }
        public SpFolder BaseDriveFolder { get; set; }
    }
}