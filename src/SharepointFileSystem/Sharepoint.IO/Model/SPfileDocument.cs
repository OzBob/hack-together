using Microsoft.Graph.Models;

namespace Sharepoint.IO
{
    public class SPfileDocument
    {
        private string? _absoluteUrlPath;
        private string? _driveId;
        private string? _fileId;

        
        public string? SpDeepLinkUrl { get { return _absoluteUrlPath; } set { _absoluteUrlPath = value; } }
        public string? DriveId { get { return _driveId; } set { _driveId = value; } }
        public string? FileId { get { return _fileId; } set { _fileId = value; } }

        public SPfileDocument()
        {
        }

        public SPfileDocument(string absoluteUrlPath)
        {
            _absoluteUrlPath = absoluteUrlPath;
            
        }

        public SPfileDocument(Uri objRec)
        {
            _absoluteUrlPath = objRec.AbsolutePath;
        }
        public SPfileDocument(Microsoft.Graph.Models.DriveItem file):this(file.WebUrl, file.ParentReference?.Id, file.Id)
        {
        }

        public SPfileDocument(string? absoluteUrlPath, string? driveId, string? fileId)
        {
            //throw nullReferenceException if any parameter is null
            if (absoluteUrlPath == null) throw new System.NullReferenceException("absoluteUrlPath");
            if (driveId == null) throw new System.NullReferenceException("driveId");
            if (fileId == null) throw new System.NullReferenceException("fileId");
            _driveId = driveId;
            _fileId = fileId;
            var uri = new Uri(absoluteUrlPath);
            var uri1 = AddParamToUriQueryString(uri, "driveId", driveId);
            var uri2 = AddParamToUriQueryString(uri1, "fileId", fileId);
            _absoluteUrlPath = uri2.AbsolutePath;
        }

        private Uri AddParamToUriQueryString(Uri uri, string paramName, string paramValue)
        {
            var query = uri.Query;
            var queryparts = query.Split('&');
            var querypart = queryparts.Where(q => q.StartsWith($"{paramName}=")).FirstOrDefault();
            if (querypart != null) return uri;

            querypart = $"{paramName}={paramValue}";
            var querypartsnew = queryparts.Where(q => !q.StartsWith($"{paramName}=")).ToList();
            querypartsnew.Add(querypart);
            var querynew = string.Join("&", querypartsnew);
            var uriBuilder = new UriBuilder(uri);
            uriBuilder.Query = querynew;
            return uriBuilder.Uri;
        }
    }


}

