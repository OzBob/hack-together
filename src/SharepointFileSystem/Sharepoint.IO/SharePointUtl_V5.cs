namespace Sharepoint.IO;

using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

public class SharePointUtl_V5
{
    private readonly GraphServiceClient _graphServiceClient;

    public SharePointUtl_V5(GraphServiceClient graphServiceClient
        )
    {
        this._graphServiceClient = graphServiceClient;
    }

    public async Task<Site?> GetSharePointSiteByIdWithDrivesAsync(string siteid)
    {
        Site? site = null;
        var sites = new List<Site>();
        var siteCollection = await _graphServiceClient
            .Sites[$"{siteid}"]
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "drives" };
            });

        if (siteCollection == null) return site;

        site = sites.FirstOrDefault();
        return site;
    }
    public async Task<Site?> FindSubsiteByNameAsync(string siteCollectionId, string subsiteName)
    {
        var sites = await _graphServiceClient
            .Sites[$"{siteCollectionId}"]
            .Sites
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "sites                                                           " };
                requestConfiguration.QueryParameters.Filter = $"displayName eq '{subsiteName}'";
            });

        // Search for the subsite by its name within the specified site collection
        //var sites = await graphClient.Sites[siteCollectionId].Sites.Filter().GetAsync();
        var site = sites?.Value?.FirstOrDefault();
        if (site == null)
        {
            Console.WriteLine($"Subsite with name '{subsiteName}' not found in site collection '{siteCollectionId}'.");
            return null;
        }

        Console.WriteLine($"Subsite with name '{subsiteName}' found in site collection '{siteCollectionId}': {site.WebUrl}");
        return site;
    }

    public async Task<Site?> GetSharePointSiteBySiteNameWithDrivesAsync(string sitename)
    {
        Site? site = null;
        var sites = new List<Site>();
        var siteCollection = await _graphServiceClient
            .Sites
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Expand = new string[] { "drives" };
            });

        if (siteCollection == null) return site;

        site = sites
            .Where(site => site.Name == sitename)
            .FirstOrDefault();
        return site;
    }

    public async Task<Drive?> GetSiteDriveIdByDriveNameAsync(Site site, string drivename)
    {
        Drive? drive = null;
        var sites = new List<Site>();
        var driveCollection = await _graphServiceClient
            .Drives
            .GetAsync();

        if (driveCollection == null || driveCollection.Value == null || driveCollection.Value.Count == 0) return drive;

        drive = driveCollection.Value.Where(d => d.Name == drivename).FirstOrDefault();
        return drive;
    }

    public async Task<DriveItem?> GetDriveFolderByFolderNameAsync(string driveid, string foldername)
    {
        DriveItem? folder = null;
        var driveRoot = await _graphServiceClient
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

    public async Task<DriveItem?> GetFolderFileByFileNameAsync(string driveid, string driveitemid, string filename)
    {
        DriveItem? file = null;
        var children = await _graphServiceClient
            .Drives[$"{driveid}"]
            .Items[driveitemid]
            .Children
            .GetAsync();
        if (children?.Value == null) return file;
        /*
        child[3](01DREI336CM6YWOERSTNDJ2YQJNQ4W6QMV):
        Name:TopDoc.docx:
        OdataType(#microsoft.graph.driveItem)
        "WebUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/Doc.aspx?sourcedoc=%7B67B167C2-3212-469B-9D62-096C396F4195%7D\\u0026file=TopDoc.docx\\u0026action=default\\u0026mobileredirect=true",
        "AdditionalData": {"@microsoft.graph.downloadUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/download.aspx?UniqueId=67b167c2-3212-469b-9d62-096c396f4195\\u0026Translate=false\\u0026tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvb3pib2Iuc2hhcmVwb2ludC5jb21AMTg1NWQ2YWEtNTQ2ZC00MjlhLThiMTQtNDQyY2FiZGYzM2NlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTY4NTM0ODQzNiIsImV4cCI6IjE2ODUzNTIwMzYiLCJlbmRwb2ludHVybCI6Ims1WDgzUkpWb0ZuSnY0VWZpNE9mY3VRNnZDcFlmSEExd2Z6L0pBT0FWeDg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjciLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IkZNaTV1YlJoTTArV1JPZUE4emMwNHc9PSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJOVEU0TlROaFpUVXRPR05rTXkwME9UWmtMVGsyTUdJdFpUVXdPV1ppTXpJM09ESXkiLCJhcHBfZGlzcGxheW5hbWUiOiJNU0dyYXBoIERhZW1vbiBDb25zb2xlIFRlc3QgQXBwIiwibmFtZWlkIjoiMGFhYzllNmYtZjJhYi00MGM0LTgzMzItZWY1ODg4NTRkOTBkQDE4NTVkNmFhLTU0NmQtNDI5YS04YjE0LTQ0MmNhYmRmMzNjZSIsInJvbGVzIjoic2hhcmVwb2ludHRlbmFudHNldHRpbmdzLnJlYWR3cml0ZS5hbGwgYWxsc2l0ZXMucmVhZCBhbGxzaXRlcy53cml0ZSBhbGxmaWxlcy53cml0ZSBhbGxwcm9maWxlcy5yZWFkIiwidHQiOiIxIiwiaXBhZGRyIjoiMjAuMTkwLjE0Mi4xNzAifQ.2CYA9bz43SbaGMM4DLQ4nuq362dqzuT6_aVHLtQiRWg\\u0026ApiVersion=2.0"},
         */
        file = children.Value.Where(f => f.Name == filename && f.FileSystemInfo != null).FirstOrDefault();
        //file = children.Value.Where(f => f.Name == filename && f.FileObject != null).FirstOrDefault();
        return file;
    }
    public async Task<DriveItem?> GetFolderSubFolderByFolderNameAsync(string driveid, string driveitemid, string foldername)
    {
        DriveItem? file = null;
        var children = await _graphServiceClient
            .Drives[$"{driveid}"]
            .Items[driveitemid]
            .Children
            .GetAsync();
        if (children?.Value == null) return file;
        //TODO filter children by item."OdataType": "#microsoft.graph.driveItem" and item.Folder is not null
        file = children.Value.Where(f => f.Name == foldername && f.Folder != null).FirstOrDefault();
        return file;
    }

    /// <summary>
    /// Permission.Application: Files.Read.All, Files.ReadWrite.All, Sites.Read.All, Sites.ReadWrite.All
    /// </summary>
    /// <param name="file"></param>
    /// <returns></returns>
    public MemoryStream GetFileAsStream(DriveItem? file)
    {
        if (file?.ParentReference?.DriveId == null) return null;

        var child = _graphServiceClient
            .Drives[$"{file.ParentReference.DriveId}"]
            .Items[file.Id]
            .GetAsync().Result;
        if (child == null || child.Content == null) return null;
        return new MemoryStream(child.Content)
            {
                Position = 0
            };
    }

    //upload new file to DriveItem
    //PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content

    public async Task UploadFileToSharePoint(Stream fileStream, string fileName, string driveid, string folderUrl)
    {
        // Create the DriveItem object for the new file
        DriveItem newItem = new DriveItem
        {
            Name = fileName,
        };

        // Create the request URL for uploading the file
        string uploadUrl = folderUrl + "/" + fileName + ":/content";

        // Upload the file using the MSGraph V4 SDK
        await _graphServiceClient
            .Drives[driveid]
            .Root
            .ItemWithPath(uploadUrl)
            .Content
            .PutAsync(fileStream);

        // Optionally, you can retrieve the uploaded file item for further processing
        // DriveItem uploadedItem = await _graphServiceClient
        //     .Drives["me"]
        //     .Root
        //     .ItemWithPath(folderUrl + "/" + fileName)
        //     .Request()
        //     .GetAsync();
    }

    public async Task<DriveItem?> CreateSubfolder(string parentDriveId,string parentDriveitemid, string subfolderName)
    {
        var subfolder = new DriveItem
        {
            Name = subfolderName,
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
            {
                { "@microsoft.graph.conflictBehavior", "rename" }
            }
        };

        var newFolder = await _graphServiceClient
            .Drives[parentDriveId]
            .Items[parentDriveitemid]
            .Children
            .PostAsync(subfolder);

        return newFolder;
    }
}
