using Microsoft.Graph.Models;
using Microsoft.Graph;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Linq;
using System.Threading.Tasks;
//using Microsoft.Graph;
//using Microsoft.Graph.Auth;

namespace dotnet_console_microsoft_graph.Experiments
{
    internal class SharePointAddFolderIfMissing
    {
        async Task<DriveItem> FindFolderCreateFolderIfMissing(GraphServiceClient graphClient
            ,string siteName = "sitename"
            ,string documentLibraryName = "documents"
            ,string subfolderName = "insol6"
            ,string tenantId = "YourTenantId"
            ,string newSubfolderName = "subfolder")
        {

            //var confidentialClientApplication = ConfidentialClientApplicationBuilder
            //    .Create(clientId)
            //    .WithClientSecret(clientSecret)
            //    .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            //    .Build();
            //var authProvider = new ClientCredentialProvider(confidentialClientApplication);
            //var graphClient = new GraphServiceClient(authProvider);
            try
            {
                // Find the site by name
                var siteFilter = $"name eq '{siteName}'";
                var site = await graphClient.Sites
                .GetAsync(requestConfiguration =>
                {
                    //Expand either analytics($expand = allTime)
                    requestConfiguration.QueryParameters.Expand = new string[] { "children, analytics($expand = allTime)" };
                    //requestConfiguration.QueryParameters.Expand = new string[] { "children, analytics($expand = allTime)" };
                    requestConfiguration.QueryParameters.Filter = siteFilter;
                });
                if (site == null) throw new NullReferenceException(siteFilter);
                if (site.Value == null) throw new NullReferenceException(siteFilter);
                if (site.Value.Count == 0) throw new NullReferenceException(siteFilter);
                var targetSite = site.Value.First();
                if (targetSite != null)
                {
                    Trace.WriteLine($"Found Site: {targetSite.Id} - {targetSite.DisplayName}");

                    // Find the document library by name
                    var drives = await graphClient.Sites[targetSite.Id].Drives.GetAsync();

                    if (drives == null) throw new NullReferenceException(siteFilter);
                    if (drives.Value == null) throw new NullReferenceException(siteFilter);
                    if (drives.Value.Count == 0) throw new NullReferenceException(siteFilter);
                    var targetDrive = site.Value.FirstOrDefault(drive => drive.Name == documentLibraryName);
                    if (targetDrive == null) throw new NullReferenceException($"Drive('{documentLibraryName}') not found");

                    if (targetDrive != null)
                    {
                        Trace.WriteLine($"Found Document Library: {targetDrive.Id} - {targetDrive.Name}");

                        // Find the subfolder within the document library
                        var rootFolder = await graphClient.Drives[targetDrive.Id].Root.GetAsync();
                        if (rootFolder == null) throw new NullReferenceException(siteFilter);

                        var subfolders = await graphClient.Drives[targetDrive.Id].Items[rootFolder.Id].Children.GetAsync();

                        if (subfolders == null) throw new NullReferenceException(siteFilter);
                        if (subfolders.Value == null) throw new NullReferenceException(siteFilter);
                        if (subfolders.Value.Count == 0) throw new NullReferenceException(siteFilter);

                        var targetSubfolder = subfolders.Value.FirstOrDefault(folder => folder.Name == subfolderName);
                        
                        if (targetSubfolder != null)
                        {
                            Trace.WriteLine($"Found Subfolder: {targetSubfolder.Id} - {targetSubfolder.Name}");
                            return targetSubfolder;
                        }
                        else
                        {
                            Trace.WriteLine($"Subfolder '{subfolderName}' not found. Creating a new subfolder...");

                            // Create the new subfolder
                            var newSubfolder = new DriveItem
                            {
                                Name = newSubfolderName,
                                Folder = new Folder(),
                            };

                            var createdSubfolder = await graphClient.Drives[targetDrive.Id]
                                .Items[rootFolder.Id]
                                .Children
                                .PostAsync(newSubfolder);
                            
                            if (createdSubfolder == null) throw new NullReferenceException($"createdSubfolder('{newSubfolderName}') not created");

                            Trace.WriteLine($"New Subfolder Created: {createdSubfolder.Id} - {createdSubfolder.Name}");

                            subfolders = await graphClient.Drives[targetDrive.Id].Items[rootFolder.Id].Children.GetAsync();

                            if (subfolders == null) throw new NullReferenceException(siteFilter);
                            if (subfolders.Value == null) throw new NullReferenceException(siteFilter);
                            if (subfolders.Value.Count == 0) throw new NullReferenceException(siteFilter);

                            targetSubfolder = subfolders.Value.FirstOrDefault(folder => folder.Name == subfolderName);
                            if (targetSubfolder == null) throw new NullReferenceException($"targetSubfolder('{subfolderName}') not found");

                            return targetSubfolder;
                        }
                    }
                    else
                    {
                        throw new Exception($"Document Library '{documentLibraryName}' not found.");
                    }
                }
                else
                {
                    throw new Exception($"Site '{siteName}' not found.");
                }
            }
            catch
            {
                throw;
            }
        }
    }
}