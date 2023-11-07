using Microsoft.Graph;
using Microsoft.Graph.Models;
using MSGraphAuth;

namespace Sharepoint.IO
{
    public interface ISharePointExtDmsClient
    {
        Task<Stream?> DownloadDocStreamByIdAsync(string webUrl);
        SpDoc GetSPdocFromUrl(string WebUrl);
        string GetUrlFromDoc(SpDoc doc);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file">stream of file</param>
        /// <param name="title">document title include extension</param>
        /// <param name="siteName"></param>
        /// <param name="subsiteName">optional</param>
        /// <param name="relativePath">must not start with slash</param>
        /// <returns></returns>
        Task<SpDoc> UploadInsolDocAsync(Stream document, string title, string siteName, string subSiteName, string relativePath);
        Task<SpFolder> GetDriveFolderDocumentsMap(string drivename, string folder);
        Task<string> SetSiteIdAsync(string siteName, string subSiteName);
        string FindSubFolderId(IList<string> subfolders, bool createIfMissing = false);
        Task DeleteFile(string WebUrl);
    }
    public class SharePointExtDmsClient : ISharePointExtDmsClient
    {
        private OAuth2ClientSecretCredentialsGrantService? oAuth2ClientCredentialsGrantService;
        private Lazy<GraphServiceClient> graphClient;
        private Lazy<SharepointHelperService> sharepointHelperService;
        private Lazy<ISharePointSiteService> sharePointSiteService;
        private readonly string sitenNameUriPart;
        private string? _siteId;
        private const string INSOL6 = "INSOL6";
        public SharePointExtDmsClient(
            string sitenNameUriPart
            , string clientId
            , string clientSecret
            , string tenantId
            , string? apiUrl
            , string baseSiteUri
            , IEnumerable<string>? scopes = null)
        {
            oAuth2ClientCredentialsGrantService = new OAuth2ClientSecretCredentialsGrantService(clientId, clientSecret, tenantId, apiUrl, null);
            this.sitenNameUriPart = sitenNameUriPart;
            graphClient = new Lazy<GraphServiceClient>(this.oAuth2ClientCredentialsGrantService.GetClientSecretClient);
            sharepointHelperService = new Lazy<SharepointHelperService>(() =>
                new SharepointHelperService(graphClient.Value, INSOL6)
            );
            sharePointSiteService = new Lazy<ISharePointSiteService>(() =>
                new SharePointSiteService(graphClient.Value, baseSiteUri)
            );
        }
        public async Task<string> SetSiteIdAsync(string siteName, string subSiteName = "")
        {
            if (string.IsNullOrEmpty(_siteId))
            {
                var sitesvc = sharePointSiteService.Value;
                var isSubSite = !string.IsNullOrEmpty(subSiteName);
                bool foundSite;
                Site? mainSite = null;
                try
                {
                    //mainSite = await sitesvc.GetSiteBySiteIdOrFullPathAsync(sharepointsiteToSearch);
                    mainSite = await sitesvc.GetSiteByNameAsync(siteName);
                    foundSite = mainSite != null;
                }
                catch
                {
                    foundSite = false;
                }
                if (mainSite == null) throw new Exception($"Site not found {siteName}");

                if (foundSite && isSubSite && (mainSite != null) && mainSite.Id != null && !string.IsNullOrEmpty(subSiteName))
                {
                    var parentSiteId = mainSite.Id;
                    mainSite = await sitesvc.GetSiteSubSiteByNameAsync(parentSiteId, subSiteName);
                    if (mainSite != null) { _siteId = mainSite.Id; }
                    else
                        _siteId = await sitesvc.GetSiteIdSubSiteAsync(parentSiteId, subSiteName);
                    //if (mainSite == null) throw new Exception($"Site not found {subSiteName}");
                }
            }
            return _siteId ?? string.Empty;
        }
        public async Task<Stream?> DownloadDocStreamByIdAsync(string webUrl)
        {
            var sitesvc = sharePointSiteService.Value;
            var doc = new SpDoc(webUrl);
            string? driveId = doc?.ParentDriveId ?? "unkown";
            string? docid = doc.Id;
            if (string.IsNullOrEmpty(driveId) || string.IsNullOrEmpty(docid))
                throw new NullReferenceException($"Missing ({SpDocConstants.PARENTDRIVEID}/{SpDocConstants.DOCID})");
            return await sitesvc.GetDownloadStream(driveId, docid);
        }
        public SpDoc GetSPdocFromUrl(string webUrl)
        {
            var doc = new SpDoc(webUrl);
            return doc;
        }

        public string GetUrlFromDoc(SpDoc doc)
        {
            return doc.ToString();
        }

        public async Task<SpDoc> UploadInsolDocAsync(Stream document, string title, string siteName, string subSiteName, string relativePath)
        {
            var sitesvc = sharePointSiteService.Value;

            //getSiteId
            await SetSiteIdAsync(siteName, subSiteName);

            var siteid = _siteId;
            if (string.IsNullOrEmpty(siteid) || siteid == "unknown") throw new Exception($"Site not found {siteName}|{subSiteName}");

            //getDriveId
            var driveId = await sitesvc.GetSiteDefaultDriveIdByName(siteid);
            if (driveId == null) throw new Exception($"Shared Document Drive for {siteName}|{subSiteName} not found!");
            var INSOL6_folderId = await sitesvc.GetSiteFolderIdByName(siteid, driveId, INSOL6);
            if (INSOL6_folderId == null) throw new Exception("INSOL6 missing");

            //upload document
            //Stream document = File.OpenRead(filepath);
            var filesize = document.Length;
            var uploadedDocument = await sharePointSiteService.Value.UploadFileToDriveFolder(
                document
                , siteid
                , driveId
                , INSOL6_folderId
                , relativePath, filesize);

            return uploadedDocument;
        }

        public async Task<SpFolder> GetDriveFolderDocumentsMap(string drivename, string folder)
        {
            //navigate to Drive
            //find INSOL6 folder by name
            //recursively find all files in folder.subfolder
            var svc = this.sharepointHelperService.Value;
            var parentSiteId = await this.SetSiteIdAsync(drivename, folder);
            var folders = await svc.GetSiteSubSiteDriveNamesAsync(parentSiteId);
            var insol6 = new SpFolder { ChildFolders = folders, Name = INSOL6 };
            return insol6;
        }

        public string FindSubFolderId(IList<string> subfolders, bool createIfMissing = false)
        {
            throw new NotImplementedException();
        }
        public async Task DeleteFile(string WebUrl)
        {
            var sitesvc = sharePointSiteService.Value;
            var doc = new SpDoc(WebUrl);
            var driveId = doc.ParentDriveId;
            await sitesvc.DeleteFile(driveId ?? "unkown", doc.Id ?? "unkown");
        }
    }
}
