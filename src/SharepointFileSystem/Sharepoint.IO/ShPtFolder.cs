using Microsoft.Graph.Models;
using System.Linq;

namespace Sharepoint.IO
{
    public class SpSite
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? WebUrl { get; set; }
        //public string? Root { get; set; }
        public string? SiteCollectionRoot { get; set; }
        public string? SiteCollectionHostname { get; set; }
        public string? SiteUrl { get; set; }
        public SpSite() {
            BaseDriveFolder = new SpFolder();
        }
        public SpSite(Site site)
        {
            Id = site.Id;
            Name = site.Name;
            WebUrl = site.WebUrl;
            //SiteCollectionRoot = site.SiteCollection.Root.BackingStore.;
            SiteCollectionHostname = site.SiteCollection?.Hostname;
            SiteUrl = site.SharepointIds?.SiteUrl;
            BaseDriveFolder = new SpFolder();
        }
        public SpFolder BaseDriveFolder { get; set; }
    }
    public class SpFolder
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? WebUrl { get; set; }
        public string? CreatedBy { get; set; }
        public DateTimeOffset? CreatedDateTime { get; set; }
        public string? Description { get; set; } 
        public string? ETag { get; set; } 
        public string? LastModifiedBy { get; set; } 
        public DateTimeOffset? LastModifiedDateTime { get; set; }
        public string? ParentReference { get; set; } 
        public SpFolder()
        {
            ChildFolders = new List<SpFolder>();
            Documents = new List<SpDoc>();
        }
        public SpFolder(Drive driveItem)
        {
            if (driveItem == null) throw new ArgumentNullException(nameof(driveItem));
            Id = driveItem.Id;
            Name = driveItem.Name;
            WebUrl = driveItem.WebUrl;
            CreatedBy = driveItem.CreatedBy?.User?.DisplayName ?? "";
            CreatedDateTime = driveItem.CreatedDateTime;
            Description = driveItem.Description;
            ETag = driveItem.ETag;
            LastModifiedBy = driveItem.LastModifiedBy?.User?.DisplayName ?? "";
            LastModifiedDateTime = driveItem.LastModifiedDateTime;
            ParentReference = $"id:{driveItem.ParentReference?.DriveId},name:{driveItem.ParentReference?.Name}";
            ChildFolders = new List<SpFolder>();
            Documents = new List<SpDoc>();
        }
        public SpFolder(DriveItem? driveItem)
        {
            if (driveItem == null) throw new ArgumentNullException(nameof(driveItem));
            Id = driveItem.Id;
            Name = driveItem.Name;
            WebUrl = driveItem.WebUrl;
            CreatedBy = driveItem.CreatedBy?.User?.DisplayName ?? "";
            CreatedDateTime = driveItem.CreatedDateTime;
            Description = driveItem.Description;
            ETag = driveItem.ETag;
            LastModifiedBy = driveItem.LastModifiedBy?.User?.DisplayName ?? "";
            LastModifiedDateTime = driveItem.LastModifiedDateTime;
            ParentReference = $"id:{driveItem.ParentReference?.DriveId},name:{driveItem.ParentReference?.Name}";
            ChildFolders = new List<SpFolder>();
            Documents = new List<SpDoc>();
        }
        public IList<SpFolder> ChildFolders { get; set; }
        public IList<SpDoc> Documents { get; set; }
        public bool HasChildFolder { get; set; } = false;
        public bool HasDocuments { get; set; } = false;
        public void AddDoc(SpDoc? spDoc)
        {
            if (spDoc == null) return;
            Documents.Add(spDoc);
            HasDocuments= true;
        }
        public void AddSubFolder(SpFolder? subFolder)
        {
            if(subFolder == null) return;
            ChildFolders.Add(subFolder);
            HasChildFolder = true;
        }
    }
    public class SpDoc
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? WebUrl { get; set; }
        public string? CreatedBy { get; set; }
        public DateTimeOffset? CreatedDateTime { get; set; }
        public string? Description { get; set; }
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
        public long? Size{ get; set; }
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
            ParentDriveId= doc.ParentReference?.DriveId;
            ParentId = doc.ParentReference?.Id;
            ParentSiteId = doc.ParentReference?.SiteId;
            Size = doc.Size;
            OdataType = doc.OdataType;
            DownloadUrl = doc.AdditionalData.ContainsKey("@microsoft.graph.downloadUrl") ? doc.AdditionalData["@microsoft.graph.downloadUrl"].ToString():"";
            WsType = GetFileExtensionFromDriveItem(doc);//file extension mimetype
        }
        private static string GetFileExtensionFromDriveItem(DriveItem? file)
        {
            if (file?.FileObject != null)
            {
                return file.FileObject.MimeType ?? "docx";
            }
            return string.Empty;
        }
    }

    public class ShPtFolder
    {
        public Site? SpSite { get; set; }
        public string FolderName { get; set; } = "";

        public bool HasChildFolder { get; set; } = false;

        public bool HasDocuments { get; set; } = false;

        public IList<ShPtFolder> ChildFolders { get; set; }

        public IList<ShPtDoc> Documents { get; set; }

        public ShPtFolder()
        {
            Documents = new List<ShPtDoc>();
            ChildFolders = new List<ShPtFolder>();
            HasChildFolder = false;
            HasDocuments = false;
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
            FileName = doc.Name ?? "unkown";
            WsType = GetFileExtensionFromDriveItem(doc);//file extension mimetype
        }
        private static string GetFileExtensionFromDriveItem(DriveItem? file)
        {
            if (file?.FileObject != null)
            {
                return file.FileObject.MimeType ?? "docx";
            }
            return string.Empty;
        }
    }
}
