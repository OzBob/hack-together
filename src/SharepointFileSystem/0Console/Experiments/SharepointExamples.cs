using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Sharepoint.IO;
using System.Text.Json;

namespace dotnet_console_microsoft_graph.Experiments;

internal static class SharepointExamples {
    public static async Task Main(GraphServiceClient betaGraphClient) {
        await GetAllSharepointSitesAsync(betaGraphClient);
    }
    private static GraphServiceClient? _graphServiceClient;
    public static async Task<string> GetAllSharepointSitesAsync(GraphServiceClient graphClient) {
        _graphServiceClient = graphClient;
        await Console.Out.WriteLineAsync("BEGIN GetAllSharepointSitesAsync");
        string siteid = String.Empty;
        try {
            //get sharepoint sites
            // get all sites
            var sites = await graphClient.Sites.GetAsync();
            //requestConfiguration => requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "mail" });
            if (sites != null && sites.Value != null) {
                foreach (var site in sites.Value) {
                    if (site == null) continue;
                    Console.WriteLine($"site({site.Id}):Name:{site.Name}:{site.DisplayName}");
                    siteid = site.Id;
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
        await Console.Out.WriteLineAsync("END GetAllSharepointSitesAsync");
        return siteid ?? String.Empty;
    }
    public static string GetSiteIdFromMSGraphSharepointSiteId(string msgraphsiteid) {
        //get middle string from "ozbob.sharepoint.com,51853ae5-8cd3-496d-960b-e509fb327822,4e67c93b-7b58-49b3-8fa0-340e3db9befd"
        if (msgraphsiteid == null) return String.Empty;
        if (msgraphsiteid.Length == 0) return String.Empty;
        if (!String.IsNullOrEmpty(msgraphsiteid) && !msgraphsiteid.Contains(',')) return msgraphsiteid;
        var sp = msgraphsiteid.Split(',');
        var siteid = sp[1];
        return siteid;
    }
    public static SpSite SpSiteItem;
    public static async Task<string> GetSharepointSiteCollectionSiteIdAsync(GraphServiceClient graphClient, string siteFullRootPath) {
        var site = await graphClient
                .Sites[$"{siteFullRootPath}"]
                .GetAsync();

        if (site == null) {
            Console.WriteLine($"No Site({siteFullRootPath}");
            return "";
        }
        SpSiteItem = new SpSite(site) {SiteCollectionRoot = siteFullRootPath };
        return site.Id ?? "";
    }
    public static async Task GetSharepointSiteAsync(GraphServiceClient graphClient, string siteid) {
        _graphServiceClient = graphClient;
        await Console.Out.WriteLineAsync($"BEGIN GetSharepointSiteAsync({siteid})");
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

                        if (drive.Name == "Documents")
                        {
                            SpSiteItem.BaseDriveFolder = new SpFolder(drive);
                        }
                        Console.WriteLine($"  drive({drive.Id}):Name:{drive.Name}:WebUrl({drive.WebUrl})");
                        //await GetDriveAsync(graphClient, drive.Id ?? "unkownid");
                        //Search Drive for doc with Ctag

                        /*
                        var ctag = "390FB120-55F7-4FF7-BCF9-9A1D089A1F97";
                        var d= await GetDocumentDriveItemByCTag(drive.Id ?? "unkownid", ctag);
                        if (d != null)
                        {
                            var jsontxt = JsonSerializer.Serialize(d);
                            Console.WriteLine($"FOUND Doc{jsontxt}");
                            }
                        else
                        {
                            Console.WriteLine($"could not find {ctag}");
                        }
                         */
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

                var siteDrives = await graphClient
                   .Sites[$"{siteid}"]
                   .Drives
                   .GetAsync(requestConfiguration => {
                       //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                   });
                //site.Drives.get
                if (siteDrives != null && siteDrives.Value != null) {
                    foreach (var drive in siteDrives.Value) {
                        if (drive == null) continue;
                        Console.WriteLine($"  sitedrive({drive.Id}):Name:{drive.Name}:WebUrl({drive.WebUrl})");

                        var siteDrive = await graphClient
                                  .Sites[$"{siteid}"]
                                  .Drives[drive.Id]
                                  .GetAsync(requestConfiguration => {
                                      //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                                      //requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                                      //requestConfiguration.QueryParameters.Expand = new string[] { "root" };
                                  });
                        var cnt = 0;
                        if (siteDrive != null) {
                            Console.WriteLine("  Drive found");
                            cnt = (siteDrive.Items == null) ? 0 : siteDrive.Items.Count;
                            Console.WriteLine($"  Drive itmes({cnt})");
                            var r = siteDrive.Root;
                            cnt = (r == null || r.Children == null) ? 0 : r.Children.Count;
                            Console.WriteLine($"  Drive root children({cnt})");
                        }
                        else { Console.WriteLine("  no Drives found"); }

                        if (siteDrive != null) {
                            var r = await graphClient
                                 .Drives[drive.Id]
                                 .Root
                                 .GetAsync(requestConfiguration => {
                                     //requestConfiguration.QueryParameters.Select = new string[] { "id", "createdDateTime", "displayName" };
                                     //requestConfiguration.QueryParameters.Expand = new string[] { "drives", "lists" };
                                     requestConfiguration.QueryParameters.Expand = new string[] { "children" };
                                 });
                            cnt = (r == null || r.Children == null) ? 0 : r.Children.Count;
                            Console.WriteLine($"  Drive root children({cnt})");


                            if (r!=null && r.Children != null) {
                                var item = r.Children[0];
                                var jsontxt2 = JsonSerializer.Serialize(item);
                                Console.WriteLine($"Item({item.Name}):"+ jsontxt2);

                                if (item !=null && item.Name == "InsolDocuments")
                                {
                                    SpSiteItem.BaseDriveFolder.AddSubFolder(new SpFolder(item));
                                    var itemid = item.Id ??"unkownDriveid";
                                    SpSiteItem.BaseDriveFolder = await GetSiteDriveItemsAsync(SpSiteItem.BaseDriveFolder, graphClient, siteid, drive.Id ?? "unkownid", itemid);
                                }
                            }
                            else { Console.WriteLine("  no Drives found"); }
                        }
                        else { Console.WriteLine("  no Drives found"); }
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
                        await GetListAsync(graphClient, siteid, list.Id ?? "unkownid");
                    }
                }
                else { Console.WriteLine("  no Lists found"); }

                if (site.Items != null) {
                    foreach (var item in site.Items) {
                        if (item == null) continue;
                        Console.WriteLine($"  item({item.Id}):Name:{item.Name}:OdataType({item.OdataType})");

                    }
                }
                else { Console.WriteLine("no site Items found"); }
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
        await Console.Out.WriteLineAsync("END GetSharepointSiteAsync");
        var jsontxt = JsonSerializer.Serialize(SpSiteItem);
        Console.WriteLine("SpSiteItem");
        Console.WriteLine($"{jsontxt}");
    }

    public static async Task<DriveItem?> GetDocumentDriveItemByCTag(string driveid, string ctagValue)
    {
        try
        {
            var children = await _graphServiceClient
               .Drives[$"{driveid}"]
               .Items
                .GetAsync(requestConfiguration =>
                {
                    //Expand either analytics($expand = allTime)
                    requestConfiguration.QueryParameters.Expand = new string[] { "children, analytics($expand = allTime)" };
                    //requestConfiguration.QueryParameters.Expand = new string[] { "children, analytics($expand = allTime)" };
                    requestConfiguration.QueryParameters.Filter = $"ctag eq '{ctagValue}'";
                });

            // Get the DriveItem from the response
            return children?.Value?.FirstOrDefault();
        }
        catch (ServiceException ex)
        {
            // Handle any errors that occurred during the request
            Console.WriteLine($"Error getting DocByCtag: {ex.Message}");
            return null;
        }
    }
    public static async Task<SpFolder> GetSiteDriveItemsAsync(SpFolder folder, GraphServiceClient graphClient,string siteid, string siteDriveid, string itemid) {
        //MSGraph ERROR: https://graph.microsoft.com/v1.0/sites/51853ae5-8cd3-496d-960b-e509fb327822/drives/$count autocorrected to "$count=" - no fix so far
        var item = await graphClient
           .Drives[siteDriveid]
           .Items[itemid]
           .GetAsync(requestConfiguration => {
               //requestConfiguration.QueryParameters.Expand = new string[] { "items" };//Parsing OData Select and Expand failed: Could not find a property named 'items' on type 'microsoft.graph.driveItem'.
           });

        if (item == null) {
            await Console.Out.WriteLineAsync("NO item found with id " + itemid);
            return folder;
        }
        Console.WriteLine($"    Item({item.Id}):Name:{item.Name}:OdataType({item.OdataType}):folderChildCount:{item.Folder?.ChildCount ?? 0}");
        return GetDriveChildren(folder, graphClient, siteDriveid, item);
    }

    private static SpFolder GetDriveChildren(SpFolder folder, GraphServiceClient graphClient, string siteDriveid, DriveItem item, int i=0) {
        if (item == null) return folder;
        
        //get Drive Children
        var children = graphClient
             .Drives[$"{siteDriveid}"]
             .Items[item.Id].Children.GetAsync();
        var prefix = new string(' ', i);
        // display all drive.List.Items
        if (children?.Result?.Value != null) {
            var childrenItems = children.Result.Value;
            foreach (var child in childrenItems) {
                Console.WriteLine($"      {prefix}child[{i}]({child.Id}):Name:{child.Name}:OdataType" +
                    $"({child.OdataType}):folderChildCount:{child.Folder?.ChildCount ?? 0}");
                Console.WriteLine(JsonSerializer.Serialize(child));

                if (child == null) continue;
                if (child.FileObject != null)
                {
                    folder.AddDoc(new SpDoc(child));
                }
                else
                {
                    folder.AddSubFolder(
                        GetDriveChildren(new SpFolder(child), graphClient , siteDriveid, child, i++)
                    );
                }
            }
        }
        else { Console.WriteLine($"    no children?.Result?.Value items found"); }
        return folder;
    }

    public static async Task GetDriveAsync(GraphServiceClient graphClient, string siteDriveid) {
        //MSGraph ERROR: https://graph.microsoft.com/v1.0/sites/51853ae5-8cd3-496d-960b-e509fb327822/drives/$count autocorrected to "$count=" - no fix so far
        var _siteDrive = await graphClient
           .Drives[siteDriveid]
           .GetAsync(requestConfiguration => {
               //requestConfiguration.QueryParameters.Expand = new string[] { "items" };});//throws oData error
               //requestConfiguration.QueryParameters.Expand = new string[] { "children" };});//throws oData error
           });
        /*

        var _siteDriveItems = await graphClient
            .Drives[siteDriveid]
            .Items
            .GetAsync();// throws The 'filter' query option must be provided.


        var _siteDriveItems1 = await graphClient
            .Drives[siteDriveid]
            .Items
        .GetAsync(requestConfiguration => {
            //requestConfiguration.QueryParameters.Select = new[] { "id", "displayName" };
            requestConfiguration.QueryParameters.Filter = "startswith(displayName, 'Documents')";
            requestConfiguration.QueryParameters.Count = true;
            requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");//set the header
        });//Expand cannot be null or empty.


        var _siteDriveItems1 = await graphClient
            .Drives[siteDriveid]
            .Items
            .GetAsync(requestConfiguration => {
                requestConfiguration.QueryParameters.Expand = new string[] { "items" };//Parsing OData Select and Expand failed: Could not find a property named 'items' on type 'microsoft.graph.driveItem'.           
                requestConfiguration.QueryParameters.Filter = "startswith(displayName, 'Documents')";
                requestConfiguration.QueryParameters.Count = true;
                requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");//set the header
            });


        */

        var _siteDriveItems = await graphClient
           .Drives[siteDriveid]
           .Root
           //.GetAsync();
           .GetAsync(requestConfiguration => {
               requestConfiguration.QueryParameters.Expand = new string[] { "children" };// "items" NA, "root" NA
           });
        //var _dItems = _siteDrive?.Root; root is null.
        var _dItems = _siteDrive?.Items;
        WriteDriveContent(graphClient, siteDriveid, _dItems);
    }

    private static void WriteDriveContent(GraphServiceClient graphClient, string siteDriveid, List<DriveItem>? _dItems) {
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

    public static async Task GetListAsync(GraphServiceClient graphClient, string siteid, string listid) {
        await Console.Out.WriteLineAsync($"    GET GetSharepointSiteList({listid})");

        var _siteLists = await graphClient
                .Sites[siteid]
                .Lists[listid]
                .GetAsync(requestConfiguration => {
                    requestConfiguration.QueryParameters.Expand = new string[] { "items" };
                });//gets Shared Documents List as odatatype #microsoft.graph.list
        //BUT AGAIN - items are blank but MSGraph Explorer shows items
        //e.g. https://graph.microsoft.com/v1.0/sites/51853ae5-8cd3-496d-960b-e509fb327822/lists/811d74b4-59ea-4edc-8e20-32d7113bc677/items/4 is a document

        /*
         "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments.parentReference.id": "505ead4d-6576-436a-a831-19275d17753c",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments.odata.etag": "9b653a64-4f45-4ba6-bd4c-2eec4ced30b6,1",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments.contentType": "Folder",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr.parentReference.id": "9b653a64-4f45-4ba6-bd4c-2eec4ced30b6",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr.odata.etag": "eaa710e6-ccf2-4a3d-a161-0054650fbe8c,2",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr.contentType": "Folder",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/TopDoc.docx.id": "4",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/TopDoc.docx.parentReference.id": "eaa710e6-ccf2-4a3d-a161-0054650fbe8c",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/TopDoc.docx.odata.etag": "67b167c2-3212-469b-9d62-096c396f4195,3",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/TopDoc.docx.contentType": "Document",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/TopDoc.docx.fields@odata.context": "https://graph.microsoft.com/v1.0/$metadata#sites('51853ae5-8cd3-496d-960b-e509fb327822')/lists('811d74b4-59ea-4edc-8e20-32d7113bc677')/items('4')/fields/$entity",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/childfldr.parentReference.id": "eaa710e6-ccf2-4a3d-a161-0054650fbe8c",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/childfldr.odata.etag": "48083a3c-d887-45e0-b430-49ff4a7689d0,1",
  "Lists(811d74b4-59ea-4edc-8e20-32d7113bc677).items.InsolDocuments/topfldr/childfldr.contentType": "Folder"
         */
        //.GetAsync(requestConfiguration => {
        //    requestConfiguration.QueryParameters.Expand = new string[] { "items", "lists" };
        ////Parsing OData Select and Expand failed: Could not find a property named 'lists' on type 'microsoft.graph.list'.
        //});
        var _dItems = _siteLists?.Items;
        if (_dItems != null) {
            foreach (Microsoft.Graph.Models.ListItem item in _dItems) {
                if (item == null) continue;
                Console.WriteLine($"    Item({item.Id}):Name:{item.Name}:OdataType({item.OdataType})");
            }
        }
        else { Console.WriteLine($"    no List({listid}) items found"); }
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

