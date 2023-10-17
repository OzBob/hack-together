using Azure.Identity;
using System.Diagnostics;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Sharepoint.IO.Model;
using System.Text.Json;
using System.Drawing.Printing;
using System.Runtime.CompilerServices;

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
        Task<SpSite> MapFullSharepointSiteAsync(string siteid);
        Task GetSiteDriveItemsAsync(SpFolder insolDocFolder, string siteDriveid, string itemid);
        Task GetDriveChildren(SpFolder parent, string siteDriveid, DriveItem item, int depth = 0);
        MemoryStream GetFileAsStream(string driveId, string fileId);
        Task<DriveItem?> GetFileAsync(string folderUrl, string fileName);
    }

    public class SharepointHelperService : ISharepointHelperService
    {
        private const string SHARED_DOCUMENTS = "Shared Documents";
        private readonly GraphServiceClient _graphServiceClient;
        private readonly string _topFolderNameMustBe;
        private readonly string _defaultDriveNameUrlEndocded;
        private const int MAXDEPTH = 2;
        private bool filterByFolderName = false;
        public SharepointHelperService(GraphServiceClient graphClient, string? topFolderNameMustBe = "", string? defaultDriveName = SHARED_DOCUMENTS)
        {
            this._graphServiceClient = graphClient;
            this._topFolderNameMustBe = topFolderNameMustBe ?? "";
            //this._defaultDriveName = plainDriveName;
            var plainDriveName = defaultDriveName ?? SHARED_DOCUMENTS;
            if (!string.IsNullOrEmpty(plainDriveName))
                this._defaultDriveNameUrlEndocded = System.Text.Encodings.Web.UrlEncoder.Default.Encode(plainDriveName);
            else
                this._defaultDriveNameUrlEndocded = "";
            filterByFolderName = !string.IsNullOrEmpty(topFolderNameMustBe);
        }

        public async Task<string> GetSharepointSiteCollectionSiteIdAsync(string siteid)
        {
            var site = await _graphServiceClient
                    .Sites[$"{siteid}"]
                    .GetAsync();

            if (site == null)
            {
                Trace.WriteLine($"No Site({siteid}");
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
            file = children.Value.Where(f => f.Name == filename && f.FileSystemInfo != null).FirstOrDefault();
            //file = children.Value.Where(f => f.Name == filename && f.FileObject != null).FirstOrDefault();
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
            files = children.Value.Where(f => f.FileSystemInfo != null).ToList();
            //files = children.Value.Where(f => f.FileObject != null).ToList();
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
        /// get file as DriveItem
        /// PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content
        /// </summary>
        /// <param name="folderUrl"></param>
        /// <param name="fileName"></param>
        /// <returns >DriveItem</returns>
        public Task<DriveItem?> GetFileAsync(string folderUrl, string fileName)
        {
            // retrieve file item
            return _graphServiceClient
                 .Drives["me"]
                 .Root
                 .ItemWithPath(folderUrl + "/" + fileName)
                 .GetAsync();
        }

        /// <summary>
        /// Permission.Application: Files.Read.All, Files.ReadWrite.All, Sites.Read.All, Sites.ReadWrite.All
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public MemoryStream GetFileAsStream(string driveId, string fileId)
        {
            if (string.IsNullOrEmpty(driveId))
                throw new ArgumentNullException(nameof(driveId));
            if (string.IsNullOrEmpty(fileId))
                throw new ArgumentNullException(nameof(fileId));

            var child = _graphServiceClient
                .Drives[driveId]
                .Items[fileId]
                .GetAsync()
                .Result;
            if (child == null)
                throw new ArgumentNullException(nameof(child));
            if (child.Content == null)
                throw new ArgumentNullException("child.Content");

            return new MemoryStream(child.Content) { Position = 0 };
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
            return this.GetFileAsStream(file.ParentReference.DriveId, file.Id);
        }

        /// <summary>
        /// upload new file to DriveItem
        /// folder
        /// PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content
        /// </summary>
        /// <param name="fileStream"></param>
        /// <param name="fileName"></param>
        /// <param name="driveid"></param>
        /// <param name="folderUrl">/drives/{drive-id}/items/{parent-id}:</param>
        /// <returns></returns>
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
                Trace.WriteLine($"Error getting DriveItem: {ex.Message}");
                return null;
            }
        }

        public async Task<SpSite> MapFullSharepointSiteAsync(string siteid)
        {
            SpSite SpSiteItem = new SpSite();
            Drive? rootDrive = null;
            try
            {
                var site = await this._graphServiceClient
                    .Sites[$"{siteid}"]
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                    });
                bool foundSharedDocumentFolder = false;
                if (site != null)
                {
                    if (site.Drives != null)
                    {
                        foreach (var drive in site.Drives)
                        {
                            if (drive == null) continue;
                            //if (drive.Name == _defaultDriveName)
                            if (drive.WebUrl.EndsWith(_defaultDriveNameUrlEndocded))
                            {
                                rootDrive = drive;
                                foundSharedDocumentFolder = true;
                            }
                        }
                    }
                }
                if (foundSharedDocumentFolder && rootDrive != null)
                {
                    var driveId = rootDrive.Id ?? "unkownid";
                    SpSiteItem.BaseDriveFolder = new SpFolder(rootDrive);
                    var siteDrive = await this._graphServiceClient
                              .Sites[$"{siteid}"]
                              .Drives[driveId]
                              .GetAsync();
                    var cnt = 0;
                    if (siteDrive != null)
                    {
                        var r = await this._graphServiceClient
                             .Drives[driveId]
                             .Root
                             .GetAsync(requestConfiguration =>
                             {
                                 requestConfiguration.QueryParameters.Expand = new string[] { "children" };
                             });
                        cnt = (r == null || r.Children == null) ? 0 : r.Children.Count;
                        Trace.WriteLine($"  Drive root children({cnt})");
                        if (r != null && r.Children != null)
                        {
                            foreach (var item in r.Children)
                            {
                                if (item == null ||  string.IsNullOrEmpty(item.Id)) continue;
                                if (
                                    (filterByFolderName && item.Name == _topFolderNameMustBe)
                                    ||
                                    !filterByFolderName
                                    )
                                {
                                    var insolDocFolder = new SpFolder(item);
                                    await GetSiteDriveItemsAsync(insolDocFolder, driveId, item.Id);
                                    SpSiteItem.BaseDriveFolder.AddSubFolder(insolDocFolder);
                                }
                            }
                        }
                        else { Trace.WriteLine("  no Drives found"); }
                    }
                }
                var siteDrives = await this._graphServiceClient
                   .Sites[$"{siteid}"]
                   .Drives
                   .GetAsync(requestConfiguration =>
                   {
                       //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                   });
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                Trace.WriteLine($"Error({ex?.Error?.Code}):{ex?.Error?.Message}");
            }
            catch (AuthenticationFailedException ex)
            {
                Trace.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }
            return SpSiteItem;
        }

        public async Task GetSiteDriveItemsAsync(SpFolder insolDocFolder, string siteDriveid, string itemid)
        {
            var item = await this._graphServiceClient
               .Drives[siteDriveid]
               .Items[itemid]
               .GetAsync();

            if (item == null) return;
            var depth = 0;
            await GetDriveChildren(insolDocFolder, siteDriveid, item, depth);
        }

        public async Task GetDriveChildren(SpFolder parent, string siteDriveid, DriveItem item, int depth = 0)
        {
            if (item == null) return;
            if (depth > MAXDEPTH) return;

            //get Drive Children
            var children = this._graphServiceClient
                 .Drives[$"{siteDriveid}"]
                 .Items[item.Id].Children.GetAsync();
            if (children?.Result?.Value != null)
            {
                foreach (var child in children.Result.Value)
                {
                    if (child == null) continue;
                    //if (child.FileObject != null)
                    if (child?.FileSystemInfo != null && child?.File != null)
                    {
                        parent.AddDoc(new SpDoc(child));
                    }
                    else
                    {
                        var subfolder = new SpFolder(child);
                        await GetDriveChildren(subfolder, siteDriveid, child, depth++);
                        parent.AddSubFolder(subfolder);
                    }
                }
            }
            return;
        }
        /*
         pages
        
        do
        {
            // Use the $top query parameter to specify the page size.
            var request = graphServiceClient.Sites[siteId].Sites.Request().Top(pageSize);
            var page = await request.GetAsync();
            
            pageSites = page.CurrentPage;
            allSites.AddRange(pageSites);

            // Continue to the next page if available.
            request = page.NextPageRequest;
        } while (pageSites.Count == pageSize);


        var usersResponse = await graphServiceClient
    .Users
    .GetAsync(requestConfiguration => { 
        requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime" }; 
        requestConfiguration.QueryParameters.Top = 1; 
        });

var userList = new List<User>();
var pageIterator = PageIterator<User,UserCollectionResponse>.CreatePageIterator(graphServiceClient,usersResponse, (user) => { userList.Add(user); return true; });

await pageIterator.IterateAsync();
         */
        public async Task<IList<SpFolder>> GetSiteSubSiteDriveNamesAsync(GraphServiceClient graphClient, string siteid)
        {
            IList<SpFolder> siteFolders = new List<SpFolder>();
            try
            {
                List<Site>? allSites = new List<Site>();
                List<Site>? pageSites = new List<Site>();
                int pageSize = 10;
                //attempt with paging
                // Use the $top query parameter to specify the page size.
                var request = graphClient.Sites[siteid].Sites;
                var page = await request.GetAsync(requestConfiguration =>
                    { requestConfiguration.QueryParameters.Top = pageSize; }
                );
                var pageIterator = PageIterator<Site, SiteCollectionResponse>
                    .CreatePageIterator(graphClient, page, (site) => { allSites.Add(site); return true; });
                await pageIterator.IterateAsync();
                    
                //foreach(var s in allSites)
                //{
                //    Debug.WriteLine(s.Name);
                //}

                //var parentsite = await _graphServiceClient
                //   .Sites[$"{siteid}"].GetAsync(requestConfiguration =>
                //   {
                //       requestConfiguration.QueryParameters.Expand = new string[] { "sites" };
                //       requestConfiguration.QueryParameters.Select = new string[] { "name, id" };
                //   });
                //var sites = parentsite.Sites.ToList();

                var siteCount = allSites.Count;
                Debug.WriteLine($"GET ALL({siteCount}) subsites under: " + siteid);
                foreach (var site in allSites)
                {
                    Debug.WriteLine(site.Name);
                    if (site != null)
                    {
                        var subsite = await graphClient
                           .Sites[$"{site.Id}"]
                           .GetAsync(requestConfiguration =>
                           {
                               requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                               //requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                           });
                        if (subsite == null || subsite.Lists == null || subsite.Lists.Count == 0) { continue; }
                        var siteLists = subsite.Lists.ToList();
                        if (subsite == null || subsite.Drives == null || subsite.Drives.Count == 0) { continue; }
                        var siteDrives = subsite.Drives.ToList();
                        if (siteDrives != null)
                        {
                            foreach (var drive in siteDrives)
                            {
                                Debug.WriteLine(drive.Name);
                                siteFolders.Add(new SpFolder(drive));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            return siteFolders;
        }
    }
}