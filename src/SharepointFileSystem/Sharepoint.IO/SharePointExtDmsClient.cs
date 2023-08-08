using Microsoft.Graph;
using Microsoft.Graph.Models;
using MSGraphAuth;
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
        SPfileDocument UploadInsolDocAsync(string filepath, string title, string className, string subclass, string clientShortName, string authorNTusername = "", string clientCchId = "");
        string FindCategory(string parent, string child, string alternateparent, string subalternateparent = "");
        ShPtFolder GetDriveFolderDocumentsMap(string drivename, string folder);
    }
    public class SharePointExtDmsClient : ISharePointExtDmsClient
    {
        private OAuth2ClientCredentialsGrantService? oAuth2ClientCredentialsGrantService;
        private Lazy<GraphServiceClient> graphClient;
        private Lazy<SharepointHelperService> sharepointHelperService;        
        private readonly string sitenNameUriPart;
        private string? _siteId;
        public SharePointExtDmsClient(
            string sitenNameUriPart
            , string? clientId
            , string? clientSecret
            , string? instance
            , string? tenant
            , string? tenantId
            , string? apiUrl
            , IEnumerable<string>? scopes = null)
        {
            oAuth2ClientCredentialsGrantService = new OAuth2ClientCredentialsGrantService(
               clientId, clientSecret, instance, tenant, tenantId, apiUrl
                , null);
            this.sitenNameUriPart = sitenNameUriPart;
            graphClient = new Lazy<GraphServiceClient>(this.oAuth2ClientCredentialsGrantService.GetClientSecretClient);
            sharepointHelperService = new Lazy<SharepointHelperService>(() =>
            {
                return new SharepointHelperService(graphClient.Value);
            });
        }
        public async Task<string> Init()
        {
            if (string.IsNullOrEmpty(_siteId))
            {
                _siteId = await sharepointHelperService.Value.GetSharepointSiteCollectionSiteIdAsync(this.sitenNameUriPart);
            }
            return _siteId;
        }
        public Stream DownloadDocStreamById(string webUrl)
        {
            //get ctag from webUrl
            //search SP for DriveItem by DriveId Ctag
            //"AdditionalData": { "@microsoft.graph.downloadUrl": "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/download.aspx?UniqueId=67b167c2-3212-469b-9d62-096c396f4195\\u0026Translate=false\\u0026tempauth=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvb3pib2Iuc2hhcmVwb2ludC5jb21AMTg1NWQ2YWEtNTQ2ZC00MjlhLThiMTQtNDQyY2FiZGYzM2NlIiwiaXNzIjoiMDAwMDAwMDMtMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwIiwibmJmIjoiMTY4NTM0ODQzNiIsImV4cCI6IjE2ODUzNTIwMzYiLCJlbmRwb2ludHVybCI6Ims1WDgzUkpWb0ZuSnY0VWZpNE9mY3VRNnZDcFlmSEExd2Z6L0pBT0FWeDg9IiwiZW5kcG9pbnR1cmxMZW5ndGgiOiIxMjciLCJpc2xvb3BiYWNrIjoiVHJ1ZSIsImNpZCI6IkZNaTV1YlJoTTArV1JPZUE4emMwNHc9PSIsInZlciI6Imhhc2hlZHByb29mdG9rZW4iLCJzaXRlaWQiOiJOVEU0TlROaFpUVXRPR05rTXkwME9UWmtMVGsyTUdJdFpUVXdPV1ppTXpJM09ESXkiLCJhcHBfZGlzcGxheW5hbWUiOiJNU0dyYXBoIERhZW1vbiBDb25zb2xlIFRlc3QgQXBwIiwibmFtZWlkIjoiMGFhYzllNmYtZjJhYi00MGM0LTgzMzItZWY1ODg4NTRkOTBkQDE4NTVkNmFhLTU0NmQtNDI5YS04YjE0LTQ0MmNhYmRmMzNjZSIsInJvbGVzIjoic2hhcmVwb2ludHRlbmFudHNldHRpbmdzLnJlYWR3cml0ZS5hbGwgYWxsc2l0ZXMucmVhZCBhbGxzaXRlcy53cml0ZSBhbGxmaWxlcy53cml0ZSBhbGxwcm9maWxlcy5yZWFkIiwidHQiOiIxIiwiaXBhZGRyIjoiMjAuMTkwLjE0Mi4xNzAifQ.2CYA9bz43SbaGMM4DLQ4nuq362dqzuT6_aVHLtQiRWg\\u0026ApiVersion=2.0"},
            throw new NotImplementedException();
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

        public SPfileDocument UploadInsolDocAsync(string filepath, string title, string className, string subclass, string clientShortName, string authorNTusername = "", string clientCchId = "")
        {

            //TODO upload document
            //set SpFileDOcuemnt(DriveItem.WebUrl)
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
    }
}
