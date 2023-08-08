using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text.Json;

namespace Sharepoint.IO
{
    public interface ISharepointHelperService
    {
        Task<Site?> GetSharePointSiteByIdWithDrivesAsync(string siteid);
        Task<DriveItem?> CreateSubfolder(string parentDriveId, string parentDriveitemid, string subfolderName);
        Task<DriveItem?> GetDriveFolderByFolderNameAsync(string driveid, string foldername);
        MemoryStream GetFileAsStream(DriveItem? file);
        Task<DriveItem?> GetFolderFileByFileNameAsync(string driveid, string driveitemid, string filename);
        Task<List<DriveItem>> GetFolderFilesAsync(string driveid, string driveitemid);
        Task<DriveItem?> GetFolderSubFolderByFolderNameAsync(string driveid, string driveitemid, string foldername);
        Task<Site?> GetSharePointSiteBySiteNameWithDrivesAsync(string sitename);
        Task<Drive?> GetSiteDriveIdByDriveNameAsync(string drivename);
        Task UploadFileToSharePoint(Stream fileStream, string fileName, string driveid, string folderUrl);
        Task<string> GetSharepointSiteCollectionSiteIdAsync(string siteid);
        Task<SpSite> GetSharepointSiteAsync(GraphServiceClient graphClient, string siteid);
        Task GetSiteDriveItemsAsync(SpFolder insolDocFolder, GraphServiceClient graphClient, string siteDriveid, string itemid);
        Task GetDriveChildren(SpFolder parent, GraphServiceClient graphClient, string siteDriveid, DriveItem item, int depth = 0);
    }

    public class SharepointHelperService : ISharepointHelperService
    {
        private readonly GraphServiceClient _graphServiceClient;
        private const int MAXDEPTH = 2;

        public SharepointHelperService(GraphServiceClient graphClient)
        {
            this._graphServiceClient = graphClient;
        }

        public async Task<string> GetSharepointSiteCollectionSiteIdAsync(string siteid)
        {
            var site = await _graphServiceClient
                    .Sites[$"{siteid}"]
                    .GetAsync();

            if (site == null)
            {
                Console.WriteLine($"No Site({siteid}");
                return "";
            }

            return site.Id ?? "";
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

        public async Task<Drive?> GetSiteDriveIdByDriveNameAsync(string drivename)
        {
            Drive? drive = null;
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
            if (driveRoot == null || driveRoot.Children == null || driveRoot.Children.Count == 0) return null;
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
            file = children.Value.Where(f => f.Name == filename && f.FileObject != null).FirstOrDefault();
            return file;
        }
        public async Task<List<DriveItem>> GetFolderFilesAsync(string driveid, string driveitemid)
        {
            List<DriveItem> files = new List<DriveItem>();

            var children = await _graphServiceClient
                .Drives[$"{driveid}"]
                .Items[driveitemid]
                .Children
                .GetAsync();
            if (children?.Value == null) return files;
            /*
            child[3](01DREI336CM6YWOERSTNDJ2YQJNQ4W6QMV):
            Name:TopDoc.docx:
            OdataType(#microsoft.graph.driveItem)
            "WebUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/Doc.aspx?sourcedoc=%7B67B167C2-3212-469B-9D62-096C396F4195%7D\\u0026file=TopDoc.docx\\u0026action=default\\u0026mobileredirect=true",
            "AdditionalData": {"@microsoft.graph.downloadUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/download.aspx?UniqueId=67b167c2-3212-469b-9d62-096c396f4195\\u0026Translate=false\\u0026tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvb3pib2Iuc2hhcmVwb2ludC5jb21AMTg1NWQ2YWEtNTQ2ZC00MjlhLThiMTQtNDQyY2FiZGYzM2NlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTY4NTM0ODQzNiIsImV4cCI6IjE2ODUzNTIwMzYiLCJlbmRwb2ludHVybCI6Ims1WDgzUkpWb0ZuSnY0VWZpNE9mY3VRNnZDcFlmSEExd2Z6L0pBT0FWeDg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjciLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IkZNaTV1YlJoTTArV1JPZUE4emMwNHc9PSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJOVEU0TlROaFpUVXRPR05rTXkwME9UWmtMVGsyTUdJdFpUVXdPV1ppTXpJM09ESXkiLCJhcHBfZGlzcGxheW5hbWUiOiJNU0dyYXBoIERhZW1vbiBDb25zb2xlIFRlc3QgQXBwIiwibmFtZWlkIjoiMGFhYzllNmYtZjJhYi00MGM0LTgzMzItZWY1ODg4NTRkOTBkQDE4NTVkNmFhLTU0NmQtNDI5YS04YjE0LTQ0MmNhYmRmMzNjZSIsInJvbGVzIjoic2hhcmVwb2ludHRlbmFudHNldHRpbmdzLnJlYWR3cml0ZS5hbGwgYWxsc2l0ZXMucmVhZCBhbGxzaXRlcy53cml0ZSBhbGxmaWxlcy53cml0ZSBhbGxwcm9maWxlcy5yZWFkIiwidHQiOiIxIiwiaXBhZGRyIjoiMjAuMTkwLjE0Mi4xNzAifQ.2CYA9bz43SbaGMM4DLQ4nuq362dqzuT6_aVHLtQiRWg\\u0026ApiVersion=2.0"},
             */
            files = children.Value.Where(f => f.FileObject != null).ToList();
            return files;
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
            if (file?.ParentReference?.DriveId == null)
                throw new ArgumentNullException(nameof(file.ParentReference.DriveId));

            var child = _graphServiceClient
                .Drives[$"{file.ParentReference.DriveId}"]
                .Items[file.Id]
                .GetAsync().Result;
            if (child == null)
                throw new ArgumentNullException(nameof(child));
            if (child.Content == null)
                throw new ArgumentNullException("child.Content");

            return new MemoryStream(child.Content) { Position = 0 };
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

        public async Task<DriveItem?> CreateSubfolder(string parentDriveId, string parentDriveitemid, string subfolderName)
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

        public async Task<DriveItem?> GetDocumentDriveItemByCTag(string driveid, string ctagValue)
        {
            try
            {
                var children = await _graphServiceClient
                   .Drives[$"{driveid}"]
                   .Items
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Expand = new string[] { "children" };
                        requestConfiguration.QueryParameters.Filter = $"ctag eq '{ctagValue}'";
                    });

                // Get the DriveItem from the response
                return children?.Value?.FirstOrDefault();
            }
            catch (ServiceException ex)
            {
                // Handle any errors that occurred during the request
                Console.WriteLine($"Error getting DriveItem: {ex.Message}");
                return null;
            }
        }
       
        public async Task<SpSite> GetSharepointSiteAsync(GraphServiceClient graphClient, string siteid)
        {
            SpSite SpSiteItem = new SpSite();
            try
            {
                var site = await graphClient
                    .Sites[$"{siteid}"]
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                    });
                if (site != null)
                {
                    if (site.Drives != null)
                    {
                        foreach (var drive in site.Drives)
                        {
                            if (drive == null) continue;
                            if (drive.Name == "Documents")
                            {
                                SpSiteItem.BaseDriveFolder = new SpFolder(drive);
                                if (drive == null) continue;
                                var siteDrive = await graphClient
                                          .Sites[$"{siteid}"]
                                          .Drives[drive.Id]
                                          .GetAsync(requestConfiguration =>
                                          {
                                          });
                                var cnt = 0;
                                if (siteDrive != null)
                                {
                                    var r = await graphClient
                                         .Drives[drive.Id]
                                         .Root
                                         .GetAsync(requestConfiguration =>
                                         {
                                             requestConfiguration.QueryParameters.Expand = new string[] { "children" };
                                         });
                                    cnt = (r == null || r.Children == null) ? 0 : r.Children.Count;
                                    Console.WriteLine($"  Drive root children({cnt})");
                                    if (r != null && r.Children != null)
                                    {
                                        var item = r.Children[0];
                                        var jsontxt2 = JsonSerializer.Serialize(item);
                                        if (item != null && item.Name == "InsolDocuments")
                                        {
                                            var insolDocFolder = new SpFolder(item);
                                            var itemid = item.Id ?? "unkownDriveid";
                                            var driveId = drive.Id ?? "unkownid";
                                            await GetSiteDriveItemsAsync(insolDocFolder, graphClient, driveId, itemid);                                            
                                            SpSiteItem.BaseDriveFolder.AddSubFolder(insolDocFolder);
                                        }
                                    }
                                    else { Console.WriteLine("  no Drives found"); }
                                }
                            }
                        }
                    }
                }

                var siteDrives = await graphClient
                   .Sites[$"{siteid}"]
                   .Drives
                   .GetAsync(requestConfiguration =>
                   {
                       //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                   });
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                Console.WriteLine($"Error({ex?.Error?.Code}):{ex?.Error?.Message}");
            }
            catch (AuthenticationFailedException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return SpSiteItem;
        }

        public async Task GetSiteDriveItemsAsync(SpFolder insolDocFolder, GraphServiceClient graphClient, string siteDriveid, string itemid)
        {
            var item = await graphClient
               .Drives[siteDriveid]
               .Items[itemid]
               .GetAsync();

            if (item == null) return;
            var depth = 0;
            await GetDriveChildren(insolDocFolder, graphClient, siteDriveid, item, depth);
        }

        public async Task GetDriveChildren(SpFolder parent, GraphServiceClient graphClient, string siteDriveid, DriveItem item, int depth = 0)
        {
            if (item == null) return;
            if (depth > MAXDEPTH) return;

            //get Drive Children
            var children = graphClient
                 .Drives[$"{siteDriveid}"]
                 .Items[item.Id].Children.GetAsync();
            if (children?.Result?.Value != null)
            {
                foreach (var child in children.Result.Value)
                {
                    if (child == null) continue;
                    if (child.FileObject != null)
                    {
                        parent.AddDoc(new SpDoc(child));
                    }
                    else
                    {
                        var subfolder = new SpFolder(child);
                        await GetDriveChildren(subfolder, graphClient, siteDriveid, child, depth++);
                        parent.AddSubFolder(subfolder);
                    }
                }
            }
            return;
        }
    }
}