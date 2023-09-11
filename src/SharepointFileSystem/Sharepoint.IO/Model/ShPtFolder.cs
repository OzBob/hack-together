using Azure;
using Microsoft.Graph.Models;
using System.Linq;

namespace Sharepoint.IO
{
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
            HasDocuments = true;
        }
        public void AddSubFolder(SpFolder? subFolder)
        {
            if (subFolder == null) return;
            ChildFolders.Add(subFolder);
            HasChildFolder = true;
        }
    }
    public class ShPtFolder
    {
        public SpSite? SpSite { get; set; }
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
}
