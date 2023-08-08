
namespace dotnet_console_microsoft_graph.Experiments
{

    using Microsoft.Graph;
    using Microsoft.Graph.Models;
    using System.Collections.Generic;
    using System.Threading.Tasks;

    public class SharePointUtl_V5
    {
        private readonly GraphServiceClient _graphServiceClient;

        public SharePointUtl_V5(GraphServiceClient graphServiceClient
            )
        {
            this._graphServiceClient = graphServiceClient;
        }

        public async Task<Site?> GetSharePointSiteByIdWithDrivesAsync(GraphServiceClient graphClient, string siteid)
        {
            Site? site = null;
            var sites = new List<Site>();
            var siteCollection = await graphClient
                .Sites[$"{siteid}"]
                .GetAsync(requestConfiguration =>
                {
                    //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                    //requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                    requestConfiguration.QueryParameters.Expand = new string[] { "drives" };
                });

            if (siteCollection == null) return site;

            site = sites.FirstOrDefault();
            return site;
        }

        public async Task<Site?> GetSharePointSiteBySiteNameWithDrivesAsync(GraphServiceClient graphClient, string sitename)
        {
            Site? site = null;
            var sites = new List<Site>();
            var siteCollection = await graphClient
                .Sites
                .GetAsync(requestConfiguration =>
                {
                    //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                    //requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                    requestConfiguration.QueryParameters.Expand = new string[] { "drives" };
                });

            if (siteCollection == null) return site;

            site = sites
                .Where(site => site.Name == sitename)
                .FirstOrDefault();
            return site;
        }

        public async Task<Drive?> GetSiteDriveIdByDriveNameAsync(GraphServiceClient graphClient, Site site, string drivename)
        {
            Drive? drive = null;
            var sites = new List<Site>();
            var driveCollection = await graphClient
                .Drives
                .GetAsync(requestConfiguration =>
                {
                    //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                    //requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                    //requestConfiguration.QueryParameters.Expand = new string[] { "drives" };
                });

            if (driveCollection == null || driveCollection.Value == null || driveCollection.Value.Count == 0) return drive;

            drive = driveCollection.Value.Where(d => d.Name == drivename).FirstOrDefault();
            return drive;
        }

        public async Task<DriveItem?> GetDriveFolderByFolderNameAsync(GraphServiceClient graphServiceClient, string driveid, string foldername)
        {
            DriveItem? folder = null;
            var driveRoot = await graphServiceClient
                    .Drives[driveid]
                    .Root
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Expand = new string[] { "children" };
                    });
            if (driveRoot == null || driveRoot.Children == null || driveRoot.Children.Count == 0) return folder;
            folder = driveRoot.Children.Where(c => c.Name == foldername).FirstOrDefault();
            return folder;
        }

        public async Task<DriveItem?> GetFolderFileByFileNameAsync(GraphServiceClient graphServiceClient, string driveid, string driveitemid, string filename)
        {
            DriveItem? file = null;
            var children = await graphServiceClient
                .Drives[$"{driveid}"]
                .Items[driveitemid]
                .Children
                .GetAsync();
            if (children?.Value == null) return file;
            //TODO filter item.FileObject is not null
            /*
            child[3](01DREI336CM6YWOERSTNDJ2YQJNQ4W6QMV):Name:TopDoc.docx:OdataType(#microsoft.graph.driveItem):folderChildCount:0
            "AdditionalData": {
		"@microsoft.graph.downloadUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/download.aspx?UniqueId=67b167c2-3212-469b-9d62-096c396f4195\\u0026Translate=false\\u0026tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvb3pib2Iuc2hhcmVwb2ludC5jb21AMTg1NWQ2YWEtNTQ2ZC00MjlhLThiMTQtNDQyY2FiZGYzM2NlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTY4NTM0ODQzNiIsImV4cCI6IjE2ODUzNTIwMzYiLCJlbmRwb2ludHVybCI6Ims1WDgzUkpWb0ZuSnY0VWZpNE9mY3VRNnZDcFlmSEExd2Z6L0pBT0FWeDg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjciLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IkZNaTV1YlJoTTArV1JPZUE4emMwNHc9PSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJOVEU0TlROaFpUVXRPR05rTXkwME9UWmtMVGsyTUdJdFpUVXdPV1ppTXpJM09ESXkiLCJhcHBfZGlzcGxheW5hbWUiOiJNU0dyYXBoIERhZW1vbiBDb25zb2xlIFRlc3QgQXBwIiwibmFtZWlkIjoiMGFhYzllNmYtZjJhYi00MGM0LTgzMzItZWY1ODg4NTRkOTBkQDE4NTVkNmFhLTU0NmQtNDI5YS04YjE0LTQ0MmNhYmRmMzNjZSIsInJvbGVzIjoic2hhcmVwb2ludHRlbmFudHNldHRpbmdzLnJlYWR3cml0ZS5hbGwgYWxsc2l0ZXMucmVhZCBhbGxzaXRlcy53cml0ZSBhbGxmaWxlcy53cml0ZSBhbGxwcm9maWxlcy5yZWFkIiwidHQiOiIxIiwiaXBhZGRyIjoiMjAuMTkwLjE0Mi4xNzAifQ.2CYA9bz43SbaGMM4DLQ4nuq362dqzuT6_aVHLtQiRWg\\u0026ApiVersion=2.0"
	},
             */
            file = children.Value.Where(f => f.Name == filename && f.FileObject != null).FirstOrDefault();
            return file;
        }
        public async Task<DriveItem?> GetFolderSubFolderByFolderNameAsync(GraphServiceClient graphServiceClient, string driveid, string driveitemid, string foldername)
        {
            DriveItem? file = null;
            var children = await graphServiceClient
                .Drives[$"{driveid}"]
                .Items[driveitemid]
                .Children
                .GetAsync();
            if (children?.Value == null) return file;
            //TODO filter children by item."OdataType": "#microsoft.graph.driveItem" and item.Folder is not null
            file = children.Value.Where(f => f.Name == foldername && f.Folder != null).FirstOrDefault();
            return file;
        }
    }
}
