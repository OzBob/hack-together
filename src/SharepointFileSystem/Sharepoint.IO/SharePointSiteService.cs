using Azure;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Models.Security;
using Microsoft.Kiota.Abstractions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint.IO
{
    public interface ISharePointSiteService
    {
        Task<Site?> GetSiteByNameAsync(string sitename);
        Task<Site?> GetSiteBySiteIdOrFullPathAsync(string siteIdOrFullPath);
        Task<Site[]> GetSites();
        Task<string?> GetSiteIdSubSiteAsync(string parentSiteId, string subSiteName);
        Task<Site?> GetSiteSubSiteByNameAsync(string parentSiteId, string subSiteName);
        Task<SpDoc> UploadFileToDriveFolder(Stream document, string siteId, string siteDriveId, string folderId, string fileName, int fileSize);
        Task<string> GetDownloadUrl(string driveId, string folderId, string fileId);
        Task<string?> GetSiteDefaultDriveIdByName(string siteId);
        Task<string?> GetSiteFolderIdByName(string siteId, string driveId, string foldername);
        /// <summary>
        /// Deletes File
        /// Throws if driveId or fileId are blank
        /// </summary>
        /// <param name="driveId">required</param>
        /// <param name="fileId">required</param>
        /// <returns></returns>
        Task DeleteFile(string driveId, string fileId);
        Task<Stream?> GetDownloadStream(string siteDriveid, string folderId, string fileId);
    }

    public class SharePointSiteService : ISharePointSiteService
    {
        private readonly string _baseSite;
        private string _baseSiteTemplate;
        private int MAXFILESIZE = 4096;
        private const string SHARED_DOCUMENTS = "Shared Documents";
        private readonly string _defaultDriveNameUrlEndocded;

        //construct a graph client GraphServiceClient
        public SharePointSiteService(GraphServiceClient graphServiceClient, string baseSite)
        {
            _graphServiceClient = graphServiceClient;
            this._baseSite = baseSite;
            this._baseSiteTemplate = $"{_baseSite}:/sites/{{0}}/";
            this._defaultDriveNameUrlEndocded = System.Text.Encodings.Web.UrlEncoder.Default.Encode(SHARED_DOCUMENTS);
        }

        public GraphServiceClient _graphServiceClient { get; }

        public async Task<string> GetDownloadUrl(string siteDriveid, string folderId, string fileId)
        {
            var result = string.Empty;
            try
            {
                var docResponses = await this._graphServiceClient
                     .Drives[$"{siteDriveid}"]
                     .Items[folderId]
                     .Children
                     .GetAsync();
                if (docResponses == null || docResponses.Value == null
                    || docResponses.Value.Count == 0) return result;
                var docResponse = docResponses.Value.Where(d => d.Id == fileId).FirstOrDefault();
                if (docResponse == null) return result;
                //var docResponse = await this._graphServiceClient
                //     .Drives[$"{siteDriveid}"]
                //     .Items[folderId]
                //     .Children[fileId]
                //     .GetAsync();
                if (docResponse == null) return result;
                var doc = new SpDoc(docResponse);
                result = doc.DownloadUrl ?? string.Empty;
            }
            catch (Exception ex)
            {
                Trace.TraceError("Failed to Find document", ex);
            }
            return result;
        }

        public Task<Stream?> GetDownloadStream(string siteDriveid, string folderId, string fileId)
        {
            /*
            var task = this._graphServiceClient
                     .Drives[$"{siteDriveid}"]
                     .Items[folderId]
                     .Children[fileId]
                     .Content
                     .GetAsync();
            */
            var task = this._graphServiceClient
                     .Drives[$"{siteDriveid}"]
                     .Items[fileId]
                     .Content
                     .GetAsync();
            return task;
        }
        public Task<Site?> GetSiteByNameAsync(string sitename)
        {
            var siteFullPath = string.Format(this._baseSiteTemplate, sitename);
            return GetSiteBySiteIdOrFullPathAsync(siteFullPath);
        }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a site by name
        public async Task<Site?> GetSiteBySiteIdOrFullPathAsync(string siteIdOrFullPath)
        {
            Trace.WriteLine($"Searching for ({siteIdOrFullPath}");
            var site = await _graphServiceClient
                 .Sites[$"{siteIdOrFullPath}"]
                    .GetAsync();
            return site;
        }

        public async Task<string?> GetSiteDefaultDriveIdByName(string siteId)
        {
            string? driveId = null;
            Drive? rootDrive = null;
            try
            {
                var site = await this._graphServiceClient
                    .Sites[$"{siteId}"]
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
                    driveId = rootDrive.Id;
                }
            }
            catch { }
            return driveId;
        }

        public async Task<string?> GetSiteFolderIdByName(string siteId, string driveId, string foldername)
        {
            string? folderId = null;
            var siteDrive = await this._graphServiceClient
                      .Sites[$"{siteId}"]
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
                        if (item == null || string.IsNullOrEmpty(item.Id)) continue;
                        if (item.Name == foldername)
                        {
                            folderId = item.Id;
                            break;
                        }
                    }
                }
            }
            return folderId;
        }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a site by name
        public async Task<Site[]> GetSites()
        {
            var siteResonse = await _graphServiceClient
                 .Sites
                 .GetAsync();

            if (siteResonse == null)
                throw new Exception($"No Sites");

            if (siteResonse.Value == null)
                throw new Exception($"No Sites value");

            if (siteResonse.Value.Count == 0)
                throw new Exception($"Zero Sites");

            return siteResonse.Value.ToArray();
        }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a subsite by name
        public async Task<string?> GetSiteIdSubSiteAsync(string parentSiteId, string subSiteName)
        {
            List<Site>? allSites = new List<Site>();
            List<Site>? pageSites = new List<Site>();
            int pageSize = 400;
            //attempt with paging
            // Use the $top query parameter to specify the page size.
            var request = _graphServiceClient.Sites[parentSiteId].Sites;
            var page = await request.GetAsync(requestConfiguration =>
                {
			        requestConfiguration.QueryParameters.Select = new[] { "id", "Name", "DisplayName" };
                    requestConfiguration.QueryParameters.Top = pageSize;
                }
            );
            var pageIterator = PageIterator<Site, SiteCollectionResponse>
                .CreatePageIterator(_graphServiceClient, page, (site) => { allSites.Add(site); return true; });
            await pageIterator.IterateAsync();

            //var site = await _graphServiceClient
            //    .Sites[$"{parentSiteId}"].GetAsync(requestConfiguration =>
            //    {
            //        requestConfiguration.QueryParameters.Expand = new string[] { "sites" };
            //    });

            //if (site == null)
            //{
            //    Trace.WriteLine($"No main Site({parentSiteId}");
            //    return null;
            //}

            //if (site.Sites == null || site.Sites.Count == 0)
            //{
            //    Trace.WriteLine($"No Subsites on Site({parentSiteId})");
            //    return null;
            //}

            var subsite = allSites
                .Where(s => s.Name == subSiteName).FirstOrDefault();
            if (subsite == null || subsite == default)
            {
                Trace.WriteLine($"No Subsite({subSiteName})");
                return null;
            }
            return subsite?.Id;
        }

        //using the Microsoft.Graph version 5 SDK
        //use the GraphServiceClient to query the Sites endpoint to find a subsite by name
        public async Task<Site?> GetSiteSubSiteByNameAsync(string parentSiteId, string subSiteName)
        {
            Site? subsite = null;
            var subSiteNameEndocded = System.Text.Encodings.Web.UrlEncoder.Default.Encode(subSiteName);
            try
            {
                /*
                 var result = await graphClient.Sites[parentSiteId].Sites.GetAsync((requestConfiguration) =>
{
    requestConfiguration.QueryParameters.Search = $"'{subSiteDisplayName}'";
});
                 */
                var request = _graphServiceClient
                    .Sites[parentSiteId]
                    .Sites
                    .GetAsync(requestConfiguration =>
                        {
                            //requestConfiguration.QueryParameters.Expand = new string[] { "sites" };
                            //Filter does not get applied!???!
                            //requestConfiguration.QueryParameters.Filter = $"Name eq '{subSiteName}'";
                            //requestConfiguration.QueryParameters.Search = $"'{subSiteName}'";
                            //requestConfiguration.QueryParameters.Search = $"\"{subSiteName}\"";
                            requestConfiguration.QueryParameters.Search = $"{subSiteName}";
                        }
                    );
                var subsiteCollection = await request;
                if (subsiteCollection != null && subsiteCollection.Value != null && subsiteCollection.Value.Count > 0)
                {
                    var subsites = subsiteCollection.Value;
                    if (subsites.Count == 1)
                        subsite = subsiteCollection.Value[0];
                    else
                    {
                        subsite = subsites.Where(ss => ss.Name == subSiteName).FirstOrDefault();
                    }
                    if (subsite == null || subsite == default)
                    {
                        Trace.WriteLine($"No Subsite({subSiteName})");
                    }
                    return subsite;
                }
            }
            catch(Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
            }
            return subsite;
        }

        public async Task<SpDoc> UploadFileToDriveFolder(Stream document, string siteId, string siteDriveId, string folderId, string fileName, int fileSize)
        {
            DriveItem? result = null;
            //call GraphClient to uload a file
            if (fileSize < MAXFILESIZE)
            {
                /*
                Upload a small file with conflictBehavior set
                To upload a small file(remember the size should not exceed 4mb according to the docs) and at the same time, set the conflictBehavior instance attribute you'll need to do it this way:
                */
                var requestInformation = _graphServiceClient
                    //.Drives[folderId]
                    //.Root
                    .Drives[siteDriveId]
                    .Items[folderId]
                    .ItemWithPath(fileName)
                    .Content
                    .ToPutRequestInformation(document);
                requestInformation.URI = new Uri(
                    requestInformation.URI.OriginalString + "?@microsoft.graph.conflictBehavior=replace");
                //requestInformation.URI.OriginalString + "?@microsoft.graph.conflictBehavior=fail");
                //requestInformation.URI.OriginalString + "?@microsoft.graph.conflictBehavior=rename");

                result = await _graphServiceClient
                    .RequestAdapter
                    .SendAsync<DriveItem>(requestInformation, DriveItem.CreateFromDiscriminatorValue);
            }
            //upload large files, the method is slightly different.
            else
            {
                var uploadSessionRequestBody = new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession
                    .CreateUploadSessionPostRequestBody
                {
                    Item = new DriveItemUploadableProperties
                    {
                        AdditionalData = new Dictionary<string, object>
                                {
                                    { "@microsoft.graph.conflictBehavior", "replace" },
                                },
                    },
                };

                // Create the upload session
                // itemPath does not need to be a path to an existing item
                var uploadSession = await _graphServiceClient
                    //.Drives[folderId]
                    //.Root
                    .Drives[siteDriveId]
                    .Items[folderId]
                    .ItemWithPath(fileName)
                    .CreateUploadSession
                    .PostAsync(uploadSessionRequestBody);

                // Max slice size must be a multiple of 320 KiB
                int maxSliceSize = 320 * 1024;
                var fileUploadTask = new LargeFileUploadTask<DriveItem>(
                    uploadSession, document, maxSliceSize, _graphServiceClient.RequestAdapter);

                var totalLength = document.Length;
                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Trace.TraceInformation($"Uploaded {prog} bytes of {totalLength} bytes");
                });
                try
                {
                    // Upload the file
                    var uploadResult = await fileUploadTask.UploadAsync(progress);
                    if (uploadResult == null || !uploadResult.UploadSucceeded)
                    {
                        throw new Exception("Upload Failed");
                    }
                    result = uploadResult.ItemResponse;
                }
                catch (ODataError ex)
                {
                    Console.WriteLine($"Error uploading: {ex.Error?.Message}");
                }
            }
            if (result == null)
                throw new Exception("Upload response empty");
            return new SpDoc(result);
        }

        public Task DeleteFile(string driveId, string fileId)
        {
            if (string.IsNullOrEmpty(driveId)) throw new ArgumentNullException("driveId");
            if (string.IsNullOrEmpty(fileId)) throw new ArgumentNullException("filedId");
            return _graphServiceClient
                   .Drives[driveId]
                   .Items[fileId]
                   .DeleteAsync();
        }
    }
}
