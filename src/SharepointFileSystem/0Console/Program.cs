using dotnet_console_microsoft_graph;
using dotnet_console_microsoft_graph.Experiments;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using MSGraphAuth;
using Sharepoint.IO;
using System.Diagnostics;
using System.IO;
using System.Text.Json;

/// This sample shows how to query the Microsoft Graph from a daemon application
/// which uses application permissions.
/// 
/// The extended project goal is to provide a Sharepoint System.IO.File and Folder abstraction for integration services
/// 

/*
 Application permissions, 
follow the instructions in .\readme.md

[X] connect to graph as daemon
[X] list AAD users
[X] list sharepoint sites
[X] list drive(s)
[X] list folders
[X] list folder subfolders
[X] list folder documents
[] perform CRUD on folder
[] perform CRUD on document
 */

var config = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json")
            .AddUserSecrets<Program>()
            .Build();

//connect to sharepoint
//var ClientAppName = config["ClientAppName"];
//var ClientAppShortName = config["ClientAppShortName"];
//var Instance = config["Instance"];
//var sharepointSiteId = config["SharepointSiteId"];
//var sharepointDriveId = config["SharepointDriveId"];
//var Tenant = config["Tenant"];
//"SiteFullRootPath":"ozbob.sharepoint.com:/sites/spfs/";
//var sharepointsiteToSearch = "ozbob.sharepoint.com:/sites/spfs/";
//try
//{
//    var _ = config["siteFullRootPath"];
//    sharepointsiteToSearch = _;
//}
//catch {
//    Console.WriteLine("noSiteSpecified");
//}
var subsiteName = "";
try
{
    var _ = config["subsiteName"];
    subsiteName = _;
}
catch
{
    Console.WriteLine("noSiteSpecified");
}
var isSubSite = !string.IsNullOrEmpty(subsiteName);
var ApiUrl = config["ApiUrl"];
var TenantId = config["TenantId"];
var ClientId = config["ClientId"];
var ClientSecret = config["ClientSecret"];
var subDocumentFolderName = config["subDocumentFolderName"];//e.g. Shared Documents
var baseSiteUri = config["SiteRootPath"];//e.g."SiteRootPath": "mcgrathnicol.sharepoint.com",
var siteName = config["SiteName"];//e.g.rstr
var subSiteName = config["SubSiteName"];//e.g.tst000005
Console.WriteLine($"search:" + siteName);
Console.WriteLine($"TenantId:" + TenantId);
//scopes are not required as this is a deamon app, and they are specified in AAD, they are listed here as a reminder to set them using the Azure Portal
var scopes = new[] {"offline_access"
    ,"SharePointTenantSettings.ReadWrite.All"
    ,"Directory.ReadWrite.All"
    ,"Sites.Read.All"
    ,"Files.ReadWrite.All"
    ,"User.Read.All"
    ,"BrowserSiteLists.ReadWrite.All"
    ,"openid", "profile", "User.Read"
    ,"analytics($expand=allTime)"
};
try
{
    var i = 0;
    Console.WriteLine(i++.ToString());
    var client = new OAuth2ClientSecretCredentialsGrantService(ClientId, ClientSecret, TenantId, ApiUrl, null);
    Console.WriteLine(i++.ToString());
    var graphClient = client.GetClientSecretClient();
    //await MSGraphExamples.ShowTenantUsersAsync(graphClient);
    //Console.WriteLine(i++.ToString());
    ISharePointSiteService sitesvc = new SharePointSiteService(graphClient, baseSiteUri);
    var userCountResponse = await graphClient.Users.Count
        .GetAsync(requestConfiguration => requestConfiguration.Headers.Add("ConsistencyLevel", "eventual"))
        //.ConfigureAwait(true).GetAwaiter().GetResult() 
        ?? 0;
    Debug.WriteLine("SP user count " + userCountResponse);

    //var siteid = await SharepointExamples.GetSharepointSiteCollectionSiteIdAsync(graphClient, sharepointsiteToSearch);
    var sites = await sitesvc.GetSites();
    var siteCount = sites.Length;
    //Console.WriteLine($"List all({siteCount}) sites START");
    //foreach (var site in sites.OrderBy(s => s.DisplayName))
    //{
    //    Console.WriteLine($"{site.DisplayName}:{site.Id}:{site.Name}:{site.WebUrl}");
    //}
    Console.WriteLine($"List all({siteCount}) sites END");

    Console.WriteLine($"Searching for ({siteName}");
    bool foundSite;
    Site? mainSite = null;
    try
    {
        //mainSite = await sitesvc.GetSiteBySiteIdOrFullPathAsync(sharepointsiteToSearch);
        mainSite = await sitesvc.GetSiteByNameAsync(siteName);
        foundSite = mainSite != null;
    }
    catch
    {
        foundSite = false;
    }
    if (mainSite == null) throw new Exception($"Site not found {siteName}");
    var siteid = "unknown";
    if (foundSite && isSubSite && (mainSite != null) && mainSite.Id != null && !string.IsNullOrEmpty(subSiteName))
    {
        var parentSiteId = mainSite.Id;
        //var svc = new Sharepoint.IO.SharepointHelperService(graphClient, subDocumentFolderName);
        //var folders = await svc.GetSiteSubSiteDriveNamesAsync(graphClient, parentSiteId);
        //foreach (var f in folders.OrderBy(fl => fl.ParentReference))
        //{
        //    Console.WriteLine($"{f.WebUrl}");
        //}
        //Console.WriteLine(i++.ToString());
        //takes too long but accurate
        mainSite = await sitesvc.GetSiteSubSiteByNameAsync(parentSiteId, subSiteName);
        if (mainSite != null) {  siteid = mainSite.Id; }
        else
            siteid = await sitesvc.GetSiteIdSubSiteAsync(parentSiteId, subSiteName);
        //if (mainSite == null) throw new Exception($"Site not found {subSiteName}");
    }
    Console.WriteLine(i++.ToString());
    if (siteid == null) throw new Exception($"SubSite not found {subSiteName}");
    //var siteid = mainSite.Id ?? "unknown";
    var runAllExamples = false;
    if (runAllExamples)
        await SharepointExamples.GetSharepointSiteAsync(graphClient, siteid);
    else
    {
        Console.WriteLine(i++.ToString());
        //var svc = new Sharepoint.IO.SharepointHelperService(graphClient, "InsolDocuments");
        var svc = new Sharepoint.IO.SharepointHelperService(graphClient, subDocumentFolderName);
        Console.WriteLine(i++.ToString());
        //var site = await svc.MapFullSharepointSiteAsync(siteid);
        //Console.WriteLine(i++.ToString());
        //var jsontxt = JsonSerializer.Serialize(site);
        //Console.WriteLine($"FOUND Doc");
        //var fileNameShPtOutput = "ShPtOutput" + DateTime.Now.ToString("yyyyMMddTHHmmssfff") + ".json";
        //File.WriteAllText(fileNameShPtOutput, jsontxt);
        var srcFileName = "testTIME.docx";
        var srcFolder = "TestDoc//";
        var newFileToUpload = srcFileName.Replace("TIME", DateTime.Now.ToString("yyyyMMddTHHmmssfff"));
        var newFileToUploadPath = srcFolder + newFileToUpload;
        //Copy "TestDoc/testTIME.docx" to 
        if (!File.Exists(newFileToUploadPath))
           File.Copy(srcFolder + srcFileName, newFileToUploadPath);

        var srcFileName2 = "testTIME2.docx";
        var newFileToUploadPath2 = srcFolder + newFileToUpload;

        Stream document = File.OpenRead(newFileToUploadPath);
        var sharePointFilePath = "1 Assets\\1.01 Circulating assets\\testsubdirB\\" + newFileToUpload;
        //var sharePointFilePath = "INSOL6//" + newFileToUpload;
        var driveId = await sitesvc.GetSiteDefaultDriveIdByName(siteid);
        if (driveId == null) throw new Exception("driveId missing");
        var INSOL6_folderId = await sitesvc.GetSiteFolderIdByName(siteid, driveId, "INSOL6");
        if (INSOL6_folderId == null) throw new Exception("folderId missing");

        var doc = await sitesvc.UploadFileToDriveFolder(
            document
            , siteid
            , driveId ?? "unkown"
            , INSOL6_folderId ?? "unkown"
            //, site.BaseDriveFolder.ChildFolders[0].Id ?? "unkown"
            , sharePointFilePath, 44);
        var docWeburl = "unkown";
        if (doc != null)
        {
            docWeburl = doc.ToString();
            Console.WriteLine("File Uploaded:" + docWeburl);
            
            var docDownloadLink = await sitesvc
                .GetDownloadUrl(driveId ?? "unkown"
                , doc.ParentId ?? "unkown"
                , doc.Id ?? "unkown");
            Console.WriteLine("File download link:" + docDownloadLink);
            if (docDownloadLink == "")
                Console.WriteLine("doc download failed");

            //upload new version
            if (File.Exists(newFileToUploadPath))
                File.Delete(newFileToUploadPath);
            File.Copy(srcFolder + srcFileName2, newFileToUploadPath);
            
            document = File.OpenRead(newFileToUploadPath);

            var doc2 = await sitesvc.UploadFileToDriveFolder(
               document
               , siteid
                , driveId ?? "unkown"
                , INSOL6_folderId ?? "unkown"
               , sharePointFilePath, 44);
            if (doc != null)
            {
                docWeburl = doc.ToString();
                Console.WriteLine("File Uploaded:" + docWeburl);
                var docDownloadLink2 = await sitesvc
                    .GetDownloadUrl(driveId ?? "unkown"
                    , doc.ParentId ?? "unkown"
                    , doc.Id ?? "unkown");
                Console.WriteLine("File download link:" + docDownloadLink2);
                var stream = await sitesvc
                    .GetDownloadStream(driveId ?? "unkown"
                    , doc.ParentId ?? "unkown"
                    , doc.Id ?? "unkown");
                if (stream != null)
                {
                    if (File.Exists("test.docx"))
                        File.Delete("test.docx");
                    using (FileStream fileStream = File.Create("test.docx"))
                    {
                        //stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(fileStream);
                    }
                    Console.WriteLine("SUCCESS dpopanw");
                }
            }

            Console.WriteLine("Press ENTER to delete");
            var key = Console.ReadLine();

            await sitesvc.DeleteFile(driveId, doc.Id);
        }
        else
        {
            Console.WriteLine("doc upload failed");
        }
    }
    var tests = new SharePointIntegrationTestsIO(graphClient);
    //Integration Tests for SharePoint FileSystem IO library

    //connect to SharePoint Documents Site

    //topfldr exists? - expect: false
    //create folder topfldr
    //topfldr exists? - expect: true

    //topfolder/childfolder exists? - expect: false
    //create child folder topfolder/childfolder
    //topfolder/childfolder exists? - expect: true
    //repeat create child folder topfolder/childfolder - expect success

    //missingFolder/missingFolder  does not exist - expect false

    //topfolder/topdoc.docx exists? expect false
    //upload new document topdoc.docx to topfldr (hint: clone TmpDoc.docx to topdoc.docx)
    //upload new document topdoc.docx to topfldr - expect success
    //topfolder/topdoc.docx exists? expect true

    //topfolder/childfolder/childoc.docx exists? expect false
    //upload new document childdoc.docx to topfolder/childfolder
    //topfolder/childfolder/childoc.docx exists? expect true
    //download topfolder/childfolder/childoc.docx to local filesystem - expect OK


    //topfolder/missingFolder  exists? - expect false
    //topfolder/missing.docx  exists? - expect false
    //download topfolder/missingFolder/missing.docx  does not exist - expect exception

}
catch (ODataError ex)
{
    if (ex.Error != null)
    {
        Console.WriteLine(ex.Error.Code);
        Console.WriteLine(ex.Error.Message);
    }
    Console.WriteLine(ex.ToString());
    if (ex.InnerException != null)
    {
        Console.WriteLine(ex.InnerException.ToString());
    }
}
catch(ServiceException sevexp)
{
    Console.WriteLine(sevexp.ToString());
    if (sevexp.InnerException != null)
    {
        Console.WriteLine(sevexp.InnerException.ToString());
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
    if (ex.InnerException != null)
    {
        Console.WriteLine(ex.InnerException.ToString());
    }
}
Console.WriteLine("TADA!");
