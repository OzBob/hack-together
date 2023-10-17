using Microsoft.Graph.Models;
using System.Diagnostics;

namespace Sharepoint.IO
{
    public class SpDoc
    {
        private const string parentidtxt = "parentid";
        private const string parentsiteidtxt = "parentsiteid";
        /// <summary>
        /// "Id": "01DREI33ZAWEHTT52V65H3Z6M2DUEJUH4X",
        /// </summary>
        public string? Id { get; set; }
        public string? Name { get; set; }
        /// <summary>
        /// e.g. "https://ozbob.sharepoint.com/sites/spfs/_layouts/15/Doc.aspx?sourcedoc=%7B390FB120-55F7-4FF7-BCF9-9A1D089A1F97%7D&file=ChildDoc.docx&action=default&mobileredirect=true",
        /// </summary>
        public string? WebUrl { get; set; }
        public string? CreatedBy { get; set; }
        public DateTimeOffset? CreatedDateTime { get; set; }
        public string? Description { get; set; }
        /// <summary>
        /// "ETag": "\"{390FB120-55F7-4FF7-BCF9-9A1D089A1F97},6\"",
        /// </summary>
        public string? ETag { get; set; }
        public string? CTag { get; set; }
        public string? LastModifiedBy { get; set; }
        public DateTimeOffset? LastModifiedDateTime { get; set; }
        public string? WsType { get; set; }
        public string? ParentPath { get; set; }
        public string? ParentDriveId { get; set; }
        public string? ParentId { get; set; }
        public string? ParentSiteId { get; set; }
        public string? DownloadUrl { get; set; }
        public long? Size { get; set; }
        public string? OdataType { get; set; }
        public SpDoc(DriveItem doc)
        {
            Id = doc.Id;
            Name = doc.Name;
            WebUrl = doc.WebUrl;
            CreatedBy = doc.CreatedBy?.User?.DisplayName ?? "";
            CreatedDateTime = doc.CreatedDateTime;
            Description = doc.Description;
            ETag = doc.ETag;
            LastModifiedBy = doc.LastModifiedBy?.User?.DisplayName ?? "";
            CTag = doc.CTag;
            LastModifiedDateTime = doc.LastModifiedDateTime;
            ParentDriveId = doc.ParentReference?.DriveId;
            ParentId = doc.ParentReference?.Id;
            ParentSiteId = doc.ParentReference?.SiteId;
            Size = doc.Size;
            OdataType = doc.OdataType;
            WsType = GetFileExtensionFromDriveItem(doc);//file extension mimetype
            DownloadUrl = doc.AdditionalData.ContainsKey("@content.downloadUrl") ? doc.AdditionalData["@content.downloadUrl"].ToString() : "";
            if (DownloadUrl == "")
                DownloadUrl = doc.AdditionalData.ContainsKey("@microsoft.graph.downloadUrl") ? doc.AdditionalData["@microsoft.graph.downloadUrl"].ToString() : "";
        }
        private static readonly char DirectorySeparatorChar = '\\';
        private static readonly char AltDirectorySeparatorChar = '/';
        private static readonly char VolumeSeparatorChar = ':';
        public static String? GetExtension(String? path)
        {
            if (path == null)
                return null;
            int length = path.Length;
            for (int i = length; --i >= 0;)
            {
                char ch = path[i];
                if (ch == '.')
                {
                    if (i != length - 1)
                        return path.Substring(i, length - i).Replace(".","");
                    else
                        return String.Empty;
                }
                if (ch == DirectorySeparatorChar || ch == AltDirectorySeparatorChar || ch == VolumeSeparatorChar)
                    break;
            }
            return String.Empty;
        }
        private static string GetFileExtensionFromDriveItem(DriveItem? file)
        {
            try
            {
                if (file?.FileSystemInfo != null && file?.File != null)
                {
                    return GetExtension(file.Name) ?? "docx";
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
                return String.Empty;
            }
            /*             
            if (file?.FileObject != null)
            {
                return file.FileObject.MimeType ?? "docx";
            }
             */
            return string.Empty;
        }
        public override string ToString()
        {
            if (string.IsNullOrEmpty(WebUrl)) return "";
            var url = WebUrl;
            var prefix = (url.IndexOf("?") > 0) ? "&" : "?";
            //if ParentId is not null or empty then add a query string parameter to the url
            if (!string.IsNullOrEmpty(ParentId))
            {
                url += $"{prefix}{parentidtxt}={ParentId}";
                prefix = "&";
            }
            if (!string.IsNullOrEmpty(ParentSiteId))
            {
                url += $"{prefix}{parentsiteidtxt}={ParentSiteId}";
                prefix = "&";
            }

            return url;
        }
    }
    public enum IDocState
    {
        Unkown,
        OK,
        MissingFromIManaage,
        MissingFromInsol,
        AddedToIManage,
        AddedToInsol,
        RemovedToIManage,
        RemovedToInsol,
        Error
    }
    public class ShPtDoc
    {
        //        DocUri = doc.ResourceLocation, //eg "https://3a8e-dmobility.imanagework-asia.com/work/link/d/AUSTRALIA!86865.1",
        //        Extension = doc.WsType,
        //        Message = doc.StatusMesage,
        //        Flags = (ulong) doc.DocState, // e.g. 0 ie Unkown, 1, OK
        //        FullCategory = folderName, // e.g. Documents
        //        Path = doc.FileName // e.g. testText File.txt
        public Uri DocUri { get; set; }

        public string DocUid { get; set; }

        public string FileName { get; set; }

        public string FullPath { get; set; }

        public IDocState DocState { get; set; } = IDocState.Unkown;

        public string? StatusMesage { get; set; }

        public string? WsType { get; set; }

        public ShPtDoc()
        {
        }

        public ShPtDoc(DriveItem doc)
            : this()
        {
            if (doc == null) return;
            var _ = new SpDoc(doc);
            FileName = _.Name ?? "unkown";
            WsType = _.WsType;//file extension mimetype
        }
    }
}
