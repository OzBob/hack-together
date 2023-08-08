using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace dotnet_console_microsoft_graph.Experiments
{
    using Microsoft.Graph;
    using System;
    using System.Threading.Tasks;
    public class SharePointUtils_V4
    {
        private readonly GraphServiceClient _graphServiceClient;

        public SharePointUtils_V4(GraphServiceClient graphServiceClient)
        {
            this._graphServiceClient = graphServiceClient;
        }



        public async Task<bool> CreateSubfolderExist(
            string siteId = "ozbob.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            string driveId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            string folderPath = "/sites/spfs/Shared Documents/InsolDocuments",
            string subfolderName = "topfldr"
        )
        {
            var folderToCreate = new DriveItem
            {
                Name = subfolderName,
                Folder = new Folder
                {
                    ChildCount = 0
                },
                AdditionalData = new Dictionary<string, object>
            {
                { "@microsoft.graph.conflictBehavior", "rename" }
            }
            };

            try
            {
                var createdFolder = await _graphServiceClient.Sites[siteId].Drives[driveId].Root.ItemWithPath(folderPath).Children
                    .Request()
                    .AddAsync(folderToCreate);

                // Return the ID of the created subfolder
                return createdFolder.Id;
            }
            catch (ServiceException ex)
            {
                // Handle exceptions
                throw;
            }

        }
        public async Task<bool> DoesSubfolderExist(
            string siteId = "ozbob.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            string driveId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            string folderPath = "/sites/spfs/Shared Documents/InsolDocuments",
            string subfolderName = "topfldr"
        )
        {
            //string siteId = "ozbob.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
            //string driveId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
            //string folderPath = "/sites/spfs/Shared Documents/InsolDocuments";
            //string subfolderName = "topfldr";

            try
            {
                var driveItem = await _graphServiceClient.Sites[siteId].Drives[driveId].Root.ItemWithPath(folderPath + "/" + subfolderName)
                    .Request()
                    .GetAsync();

                // If the subfolder exists, it will not throw an exception
                return true;
            }
            catch (ServiceException ex)
            {
                if (ex.ResponseStatusCode == (int)System.Net.HttpStatusCode.NotFound)
                {
                    // Subfolder not found
                    return false;
                }
                else
                {
                    // Handle other exceptions
                    throw;
                }
            }
        }

        public async Task<bool> CheckIfSubfolderExists1(string sharePointUrl, string parentFolderName, string subfolderName)
        {
            try
            {
                // Construct the SharePoint site URL
                string siteUrl = $"{sharePointUrl}/_api/web";

                // Retrieve the parent folder
                DriveItem parentFolder = await _graphServiceClient.Sites[siteUrl].Drive.Root.ItemWithPath(parentFolderName).Request().GetAsync();

                // Retrieve the subfolders of the parent folder
                IDriveItemChildrenCollectionPage children = await _graphServiceClient.Sites[siteUrl].Drive.Items[parentFolder.Id].Children.Request().GetAsync();

                // Check if the subfolder exists
                foreach (DriveItem childFolder in children)
                {
                    if (childFolder.Name.Equals(subfolderName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true; // Subfolder found
                    }
                }

                return false; // Subfolder not found
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during the request
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }
        public async Task<bool> DoesDocumentExist()
        {
            string siteId = "ozbob.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
            string driveId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
            string folderPath = "/sites/spfs/Shared Documents/InsolDocuments/topfldr/childfldr";
            string documentName = "done.docx";

            try
            {
                var driveItem = await _graphServiceClient.Sites[siteId].Drives[driveId].Root.ItemWithPath(folderPath + "/" + documentName)
                    .Request()
                    .GetAsync();

                // If the document exists, it will not throw an exception
                return true;
            }
            catch (ServiceException ex)
            {
                if (ex.ResponseStatusCode == (int)System.Net.HttpStatusCode.NotFound)
                {
                    // Document not found
                    return false;
                }
                else
                {
                    // Handle other exceptions
                    throw;
                }
            }
        }

    }
}
