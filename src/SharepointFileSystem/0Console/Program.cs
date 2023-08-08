using dotnet_console_microsoft_graph;
using dotnet_console_microsoft_graph.Experiments;
using Microsoft.Extensions.Configuration;
using MSGraphAuth;
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
[] list folders
[] list folder subfolders
[] list folder documents
[] perform CRUD on folder
[] perform CRUD on document
 */

var config = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json")
            .AddUserSecrets<Program>()
            .Build();

//connect to sharepoint
var ClientAppName = config["ClientAppName"];
var ClientAppShortName = config["ClientAppShortName"];
var Instance = config["Instance"];
var ApiUrl = config["ApiUrl"];
var Tenant = config["Tenant"];
var TenantId = config["TenantId"];
var ClientId = config["ClientId"];
var ClientSecret = config["ClientSecret"];
var sharepointSiteId = config["SharepointSiteId"];
var sharepointDriveId = config["SharepointDriveId"];

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
    var client = new OAuth2ClientCredentialsGrantService(
        ClientId, ClientSecret, Instance, Tenant, TenantId, ApiUrl
        , null);
    var graphClient = client.GetClientSecretClient();
    await MSGraphExamples.ShowTenantUsersAsync(graphClient);
    //sharepointSiteId = await SharepointExamples.GetAllSharepointSitesAsync(graphClient);
    var siteid = await SharepointExamples.GetSharepointSiteCollectionSiteIdAsync(graphClient, "ozbob.sharepoint.com:/sites/spfs/");

    var runAllExamples = false;
    if (runAllExamples)
        await SharepointExamples.GetSharepointSiteAsync(graphClient, siteid);
    else
    {
        var svc = new Sharepoint.IO.SharepointHelperService(graphClient);
        var site = svc.GetSharepointSiteAsync(graphClient, siteid).ConfigureAwait(true).GetAwaiter().GetResult();
        var jsontxt = JsonSerializer.Serialize(site);
        Console.WriteLine($"FOUND Doc{jsontxt}");
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
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}

