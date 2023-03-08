using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace dotnet_console_microsoft_graph.Experiments;

internal static class SharepointExamples {
    public static async Task Main(GraphServiceClient betaGraphClient) {
        await GetAllSharepointSitesAsync(betaGraphClient);
    }

    public static async Task GetAllSharepointSitesAsync(GraphServiceClient graphClient) {
        try {
            //get sharepoint sites

            // get all sites
            var sites = await graphClient.Sites.GetAsync();
            //requestConfiguration => requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "mail" });
            if (sites != null && sites.Value != null) {
                foreach (var site in sites.Value) {
                    if (site == null) continue;
                    Console.WriteLine($"site({site.Id}):Name:{site.Name}:{site.DisplayName}");
                }
                if (sites.Value.Count == 0) { Console.WriteLine("no sites found"); }
            }
            else { Console.WriteLine("no sites found"); }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) {
            Console.WriteLine($"Error({ex?.Error?.Code}):{ex?.Error?.Message}");
        }
        catch (AuthenticationFailedException ex) {
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex) {
            Console.WriteLine(ex.Message);
        }

    }
    public static async Task GetSharepointSiteAsync(GraphServiceClient graphClient, string siteid, string driveid) {
        try {
            var site = await graphClient
                .Sites[$"{siteid}"]
                .GetAsync(requestConfiguration => {
                    //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                    requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                });
            if (site != null) {
                Console.WriteLine($"site({site.Id}):Name:{site.Name}:{site.DisplayName}");

                //site.Drives.get
                if (site.Drives != null) {
                    foreach (var drive in site.Drives) {
                        if (drive == null) continue;
                        Console.WriteLine($"  drive({drive.Id}):Name:{drive.Name}:WebUrl({drive.WebUrl})");
                        await GetDriveAsync(graphClient, drive.Id ?? "unkownid");
                    }
                    //var _d = await graphClient
                    //    .Drives[driveid]
                    //    .GetAsync();
                    //.GetAsync(requestConfiguration => {
                    //requestConfiguration.QueryParameters.Expand = new string[] { "items" };});//throws oData error
                    //requestConfiguration.QueryParameters.Expand = new string[] { "children" };});//throws oData error
                    //NB all query data must be URL Encoded parameters

                    //List children from a site's drive.                    
                    // Get the site's driveId
                    /*
                    var _siteDrives = await graphClient.Sites[siteid].Drives.GetAsync();
                if (_siteDrives != null && _siteDrives.Value != null) {
                    var _drives = _siteDrives.Value;
                    // List children in the drive
                    var _drives = await graphClient.Drives[siteDriveId].GetAsync();

                    //m365 developer portal: https://developer.microsoft.com/en-us/microsoft-365/profile
                    // code examples: https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/feature/5.0/docs/upgrade-to-v5.md#drive-item-paths
                    // try https://developer.microsoft.com/en-us/graph/graph-explorer
                    // spfs site: what is site id?
                    // https://ozbob.sharepoint.com/sites/spfs/Shared Documents/Forms/AllItems.aspx?id=%2Fsites%2Fspfs%2FShared Documents%2FInsolDocuments%2Ftopfldr&viewid=415aefb0-2377-416e-b149-cd8289d4fa7e
                    // developer sharepoint site: https://ozbob.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx?viewpath=%2FShared%20Documents%2FForms%2FAllItems%2Easpx&id=%2FShared%20Documents%2Fchildfldr0&viewid=02533ae6%2Dc06a%2D4d8f%2Db48c%2D32ded2302ef4
                    // see if anyone answered: https://github.com/microsoft/hack-together/discussions/32
                    // example REST code: https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/7a2be45d2cf37f18a32cc9a60d0edf441fd23a08/docs/v4-reference-docs/driveitem-list-versions.md
                    var _dlist = await graphClient.Drives[siteDriveId].List.GetAsync();//.Items//Items[driveid].Children.GetAsync();
                    //var _ditems = await graphClient.Drives[siteDriveId].List.GetAsync();//.Items//Items[driveid].Children.GetAsync();
                    // display all drive.Items  
                    foreach (var _drive in _drives) {
                    if (_drive != null)
                    await GetDriveAsync(graphClient, _drive.Id??"unkownid");
                }
                }
                else { Console.WriteLine("    no Drive() found"); }
                     */
                }
                else { Console.WriteLine("  no Drives found"); }

                if (site.Lists != null) {
                    foreach (var list in site.Lists) {
                        if (list == null) continue;
                        Console.WriteLine($"  list({list.Id}):Name:{list.Name}:ParentReference({list.ParentReference})");
                        if (list.Items != null) {
                            foreach (var item in list.Items) {
                                if (item == null) continue;
                                Console.WriteLine($"    item({item.Id}):Name:{item.Name}:ParentReference({item.ParentReference})");
                            }
                        }
                    }
                }
                else { Console.WriteLine("  no Lists found"); }

                if (site.Items != null) {
                    foreach (var item in site.Items) {
                        if (item == null) continue;
                        Console.WriteLine($"  item({item.Id}):Name:{item.Name}:OdataType({item.OdataType})");

                    }
                }
                else { Console.WriteLine("no Items found"); }
            }
            else { Console.WriteLine("no sites found"); }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) {
            Console.WriteLine($"Error({ex?.Error?.Code}):{ex?.Error?.Message}");
        }
        catch (AuthenticationFailedException ex) {
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex) {
            Console.WriteLine(ex.Message);
        }
    }

    public static async Task GetDriveAsync(GraphServiceClient graphClient, string siteDriveid) {
        var _siteDrive= await graphClient
            .Drives[siteDriveid]
            .GetAsync(requestConfiguration => {
                //requestConfiguration.QueryParameters.Expand = new string[] { "items" };});//throws oData error
                //requestConfiguration.QueryParameters.Expand = new string[] { "children" };});//throws oData error
            });
        
        var dItems = _siteDrive?.Items;
        
        var _siteDriveItems = await graphClient
            .Drives[siteDriveid]
            .List
            .GetAsync();// throws The 'filter' query option must be provided.
        var _dItems = _siteDrive?.Items;
        if (_dItems != null) {
            foreach (Microsoft.Graph.Models.DriveItem item in _dItems) {
                if (item == null) continue;
                Console.WriteLine($"    Item({item.Id}):Name:{item.Name}:OdataType({item.OdataType}):folderChildCount:{item.Folder?.ChildCount ?? 0}");

                //get Drive Children
                var children = graphClient
                     .Drives[$"{siteDriveid}"]
                     .Items[item.Id].Children.GetAsync();
                // display all drive.List.Items
                if (children?.Result?.Value != null) {
                    var childrenItems = children.Result.Value;
                    foreach (var child in childrenItems) {
                        if (child == null) continue;
                        Console.WriteLine($"    ListItem({child.Id}):Name:{item.Name}:OdataType({item.OdataType})");
                    }
                }
                else { Console.WriteLine($"    no children?.Result?.Value items found"); }
            }
        }
        else { Console.WriteLine($"    no drive({siteDriveid}) items found"); }
    }

    public static async Task CreateNewSubDirectoryAsync(GraphServiceClient graphClient, string siteid, string driveid) {
        //v4 Get the requestInformation to make a POST request
        //var directoryObject = new DirectoryObject() {
        //    Id = Guid.NewGuid().ToString()
        //var requestInformation = graphServiceClient
        //                            .DirectoryObjects
        //                            .ToPostRequestInformation(directoryObject);
        //TODO v5 a new Sub Directory
    }
}

