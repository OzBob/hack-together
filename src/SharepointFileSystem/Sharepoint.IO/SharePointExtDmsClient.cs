using Microsoft.Graph;
using Microsoft.Graph.Models;
using MSGraphAuth;
using Sharepoint.IO.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint.IO
{
    public interface ISharePointExtDmsClient {
        Stream DownloadDocStreamById(string WebUrl);
        SPfileDocument GetSPdocFromUrl(string WebUrl);
        string GetUrlFromDoc(SPfileDocument doc);
        Task<SPfileDocument> UploadInsolDocAsync(string filepath, string title, string className, string subclass, string clientShortName, string authorNTusername = "", string clientCchId = "");
        string FindCategory(string parent, string child, string alternateparent, string subalternateparent = "");
        ShPtFolder GetDriveFolderDocumentsMap(string drivename, string folder);
        Task<string> InitAsync();
        string FindSubFolderId(IList<string> subfolders, bool createIfMissing = false);
    }
    public class SharePointExtDmsClient : ISharePointExtDmsClient
    {
        private OAuth2ClientSecretCredentialsGrantService? oAuth2ClientCredentialsGrantService;
        private Lazy<GraphServiceClient> graphClient;
        private Lazy<SharepointHelperService> sharepointHelperService;        
        private readonly string sitenNameUriPart;
        private string? _siteId;
        public SharePointExtDmsClient(
            string sitenNameUriPart
            , string clientId
            , string clientSecret
            , string tenantId
            , string? apiUrl
            , IEnumerable<string>? scopes = null)
        {
            oAuth2ClientCredentialsGrantService = new OAuth2ClientSecretCredentialsGrantService(clientId, clientSecret, tenantId, apiUrl, null);
            this.sitenNameUriPart = sitenNameUriPart;
            graphClient = new Lazy<GraphServiceClient>(this.oAuth2ClientCredentialsGrantService.GetClientSecretClient);
            sharepointHelperService = new Lazy<SharepointHelperService>(() =>
            {
                return new SharepointHelperService(graphClient.Value);
            });
        }
        public async Task<string> InitAsync()
        {
            if (string.IsNullOrEmpty(_siteId))
            {
                _siteId = await sharepointHelperService.Value.GetSharepointSiteCollectionSiteIdAsync(this.sitenNameUriPart);
            }
            return _siteId;
        }
        public Stream DownloadDocStreamById(string webUrl)
        {
            var webUrlUri = new Uri(webUrl);
            //get driveItem by driveId and fileId from webUrlUri query string
            var querystring = webUrlUri.Query;
            var querystringparts = querystring.Split('&');
            var driveId = querystringparts.Where(q => q.StartsWith("driveId=")).FirstOrDefault()?.Split('=')[1];
            //throw if driveId is null
            driveId = driveId ?? throw new ArgumentNullException(nameof(driveId));
            var fileId = querystringparts.Where(q => q.StartsWith("fileId=")).FirstOrDefault()?.Split('=')[1];
            //throw if fildId is null
            fileId = fileId ?? throw new ArgumentNullException(nameof(fileId));
            return sharepointHelperService.Value.GetFileAsStream(driveId, fileId);

        }
        public SPfileDocument GetSPdocFromUrl(string webUrl)
        {
            //get doc from sharepoint by url
            //example of weburl: "WebUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/Doc.aspx?sourcedoc=%7B67B167C2-3212-469B-9D62-096C396F4195%7D\\u0026file=TopDoc.docx\\u0026action=default\\u0026mobileredirect=true",         
            /*decoded weburl: sourcedoc={67B167C2-3212-469B-9D62-096C396F4195}\u0026file=TopDoc.docx\u0026action=default\u0026mobileredirect=true

            properties, split on ("\u0026"): 
                file=TopDoc.docx
                action=default
                mobileredirect=true
                sourcedoc={67B167C2-3212-469B-9D62-096C396F4195}
                Guid g = Guid.Parse("{67B167C2-3212-469B-9D62-096C396F4195}");
                string ctagName = g.ToString("D");
            ? can the sourcedoc GUID be used in sharepoint .net msgrpah sdk to retrieve a document?
            */

            //return SPfileDocument
            throw new NotImplementedException();
        }

        public string GetUrlFromDoc(SPfileDocument doc)
        {
            return doc.SpDeepLinkUrl ?? "";
        }

        public async Task<SPfileDocument> UploadInsolDocAsync(Stream filestream, string title, string parentFolderName, string subfolderName)
        {

            //get driveid from parentFolderName

            var driveid = "unkown";//todo get driveid
            var folderUrl = "unkown";//todo get folderUrl
            //upload document
            _siteId = await sharepointHelperService.Value.UploadFileToSharePoint(
                filestream, title, driveid, folderUrl
                );

            //set SPfileDocument(new ShPtDoc(DriveItem))
            //add query string to weburl
            //driveid
            //driveitemid
            //confirm ctag value is in weburl
            //example original weburl:"https://ozbob.sharepoint.com/sites/spfs/_layouts/15/Doc.aspx?sourcedoc=%7B67B167C2-3212-469B-9D62-096C396F4195%7D\\u0026file=TopDoc.docx\\u0026action=default\\u0026mobileredirect=true",         
            throw new NotImplementedException();
        }


        public string FindCategory(string parent, string child, string alternateparent, string subalternateparent = "")
        {
            throw new NotImplementedException();
        }

        public ShPtFolder GetDriveFolderDocumentsMap(string drivename, string folder)
        {
            //navigate to Drive
            //find folder by name
            //recursively find all files in folder.subfolder
            throw new NotImplementedException();
        }

        public string FindSubFolderId(IList<string> subfolders, bool createIfMissing = false)
        {
            throw new NotImplementedException();
        }
    }
}
