using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace dotnet_console_microsoft_graph.Experiments;

internal static class MSGraphExamples {
	public static async Task Main(GraphServiceClient betaGraphClient) {
		//// Other TokenCredentials examples are available at https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/docs/tokencredentials.md
		//var scopes = new[] { "User.Read", "Mail.Read", "User.ReadBasic.All" };
		//var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions {
		//    ClientId = "CLIENT_ID"
		//};
		//var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

		// GraphServiceClient constructor accepts tokenCredential
		//var v1GraphClient = new GraphServiceClient(tokenCredential, scopes);// client for the v1.0 endpoint
		//var betaGraphClient = new Microsoft.Graph.GraphServiceClient(tokenCredential, scopes);// client for the beta endpoint

		// Perform batch request using the beta client
		await PerformRequestWithHeaderAndQueryRequestAsync(betaGraphClient);

		// Perform batch request using the beta client
		await PerformCustomRequestWithHeaderAndQueryAsync(betaGraphClient);

		// Perform batch request using the v1 client
		await PerformBatchRequestAsync(betaGraphClient);

		// Perform paged request using the v1 client
		await IteratePagedDataAsync(betaGraphClient);
	}
	public static async Task ShowTenantUsersAsync(GraphServiceClient graphClient) {
		try {
			// Get the requestInformation to make a GET request
			var requestInformation = graphClient
									 .DirectoryObjects
									 .ToGetRequestInformation();
			Console.WriteLine("requestInformation.URI=" + requestInformation.URI);

			// get all users on tenant
			var users = await graphClient.Users.GetAsync(
				requestConfiguration => requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "mail" });
			if (users != null && users.Value != null) {
				foreach (var user in users.Value) {
					if (user == null) continue;
					Console.WriteLine($"User({user.Id}):Name:{user.DisplayName}:{user.Mail}");
				}
			}
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

	private static async Task PerformBatchRequestAsync(GraphServiceClient graphClient) {
		Console.WriteLine("-----------Performing batch requests-----------");
		var userRequest = graphClient.Me.ToGetRequestInformation();// create request object to get user information
		var messagesRequest = graphClient.Me.Messages.ToGetRequestInformation();// create request object to get user messages

		// Build the batch
		var batchRequestContent = new BatchRequestContent(graphClient);
		var userRequestId = await batchRequestContent.AddBatchRequestStepAsync(userRequest);
		var messagesRequestId = await batchRequestContent.AddBatchRequestStepAsync(messagesRequest);

		// Send the batch
		var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);

		// Get the user info
		var user = await batchResponse.GetResponseByIdAsync<User>(userRequestId);
		Console.WriteLine($"Fetched user with name {user.DisplayName} via batch");

		// Get the messages data
		var messagesResponse = await batchResponse.GetResponseByIdAsync<MessageCollectionResponse>(messagesRequestId);
		List<Message> messages = messagesResponse.Value;
		Console.WriteLine($"Fetched {messages.Count} messages via batch");
		Console.WriteLine("-----------Done with batch requests-----------");
	}

	private static async Task IteratePagedDataAsync(GraphServiceClient graphClient) {
		Console.WriteLine("-----------Performing paged requests-----------");
		var firstPage = await graphClient.Me.Messages.GetAsync();// fetch first paged of messages

		var messagesCollected = new List<Message>();
		// Build the pageIterator
		var pageIterator = PageIterator<Message, MessageCollectionResponse>.CreatePageIterator(
			graphClient,
			firstPage,
			message => {
				messagesCollected.Add(message);
				return true;
			},// per item callback
			request => {
				Console.WriteLine($"Requesting new page with url {request.URI.OriginalString}");
				return request;
			}// per request/page callback to reconfigure the request
		);

		// iterated
		await pageIterator.IterateAsync();

		// Get the messages data;
		Console.WriteLine($"Fetched {messagesCollected.Count} messages via page iterator");
		Console.WriteLine("-----------Done with paged requests-----------");
	}

	private static async Task PerformRequestWithHeaderAndQueryRequestAsync(Microsoft.Graph.GraphServiceClient graphClient) {
		Console.WriteLine("-----------Performing configured requests-----------");

		var userResponse = await graphClient.Users.GetAsync(requestConfiguration => {
			requestConfiguration.QueryParameters.Select = new[] { "id", "displayName" };// set select
			requestConfiguration.QueryParameters.Filter = "startswith(displayName, 'al')";// set filter for users displayName starting with 'al'
			requestConfiguration.QueryParameters.Count = true;
			requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");//set the header
		});

		Console.WriteLine($"Fetched {userResponse.Value.Count} users with displayName starting with 'al'");
		Console.WriteLine("-----------Done with configured requests-----------");
	}

	private static async Task PerformCustomRequestWithHeaderAndQueryAsync(Microsoft.Graph.GraphServiceClient graphClient) {
		Console.WriteLine("-----------Performing customized request-----------");

		var requestInformation = graphClient.Users.ToGetRequestInformation(requestConfiguration => {
			requestConfiguration.QueryParameters.Select = new[] { "id", "displayName" };// set select
			requestConfiguration.QueryParameters.Filter = "startswith(displayName, 'al')";// set filter for users displayName starting with 'al'
			requestConfiguration.QueryParameters.Count = true;
			requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");//set the header
		});

		var userResponse = await graphClient.RequestAdapter.SendAsync<Microsoft.Graph.Models.UserCollectionResponse>(
				requestInformation, Microsoft.Graph.Models.UserCollectionResponse.CreateFromDiscriminatorValue);

		Console.WriteLine($"Fetched {userResponse.Value.Count} users with displayName starting with 'al'");
		Console.WriteLine("-----------Done with customized requests-----------");
	}
	private static void sample() {
		/*
		 List children in the root of the current user's drive
		 GraphServiceClient graphClient = new GraphServiceClient( authProvider );
		var children = await graphClient.Me.Drive.Root.Children
		.Request()
		.GetAsync();
		
		 
		 List children of a DriveItem with a known ID
GraphServiceClient graphClient = new GraphServiceClient( authProvider );

var children = await graphClient.Drives["{drive-id}"].Items["{driveItem-id}"].Children
	.Request()
	.GetAsync();

		List children of a DriveItem with a known path
GET /drives/{drive-id}/root:/{path-relative-to-root}:/children

		OData Guidance
		https://learn.microsoft.com/en-us/graph/query-parameters?tabs=http
		Use query parameters to customize responses
		 n the beta endpoint, the $ prefix is optional. For example, instead of $filter, you can use filter. On the v1 endpoint, the $ prefix is optional for only a subset of APIs. For simplicity, always include $ if using the v1 endpoint.
		OData system query options
A Microsoft Graph API operation might support one or more of the following OData system query options. These query options are compatible with the OData V4 query language and are supported in only GET operations.


Name	Description	Example
$count	Retrieves the total count of matching resources.	/me/messages?$top=2&$count=true
$expand	Retrieves related resources.	/groups?$expand=members
$filter	Filters results (rows).	/users?$filter=startswith(givenName,'J')
$format	Returns the results in the specified media format.	/users?$format=json
$orderby	Orders results.	/users?$orderby=displayName desc
$search	Returns results based on search criteria.	/me/messages?$search=pizza
$select	Filters properties (columns).	/users?$select=givenName,surname
$skip	Indexes into a result set. Also used by some APIs to implement paging and can be used together with $top to manually page results.	/me/messages?$skip=11
$top	Sets the page size of results.	/users?$top=2
To know the OData system query options that an API and its properties support, see the Properties table in the resource page, and the Optional query parameters section of the LIST and GET operations for the API.

Other query parameters
Name	Description	Example
$skipToken	Retrieves the next page of results from result sets that span multiple pages. (Some APIs use $skip instead.)	/users?$skiptoken=X%274453707402000100000017...
Other OData URL capabilities
The following OData 4.0 capabilities are URL segments, not query parameters.

Name	Description	Example
$count	Retrieves the integer total of the collection.	GET /users/$count
GET /groups/{id}/members/$count
$ref	Updates entities membership to a collection.	POST /groups/{id}/members/$ref
$value	Retrieves or updates the binary value of an item.	GET /me/photo/$value
$batch	Combine multiple HTTP requests into a batch request.	POST /$batch
Encoding query parameters
The values of query parameters should be percent-encoded as per RFC 3986.


		http://docs.oasis-open.org/odata/odata/v4.01/odata-v4.01-part2-url-conventions.html#_Toc31360955

		5.1 System Query Options
System query options are query string parameters that control the amount and order of the data returned for the resource identified by the URL. The names of all system query options are optionally prefixed with a dollar ($) character. 4.01 Services MUST support case-insensitive system query option names specified with or without the $ prefix. Clients that want to work with 4.0 services MUST use lower case names and specify the $ prefix.

For GET, PATCH, and PUT requests the following rules apply:

·       Resource paths identifying a single entity, a complex type instance, a collection of entities, or a collection of complex type instances allow $compute, $expand and $select.

·       Resource paths identifying a collection allow $filter, $search, $count, $orderby, $skip, and $top.

·       Resource paths ending in /$count allow $filter and $search.

·       Resource paths not ending in /$count or /$batch allow $format.

For POST requests to an action URL the return type of the action determines the applicable system query options that a service MAY support, following the same rules as GET requests.

POST requests to an entity set follow the same rules as GET requests that return a single entity.

System query options SHOULD NOT be applied to a DELETE request.

An OData service may support some or all of the system query options defined. If a data service does not support a system query option, it MUST reject any request that contains the unsupported option.

The same system query option, irrespective of casing or whether or not it is prefixed with a $, MUST NOT be specified more than once for any resource.

The semantics of all system query options are defined in the [OData-Protocol] document.

The grammar and syntax rules for system query options are defined in [OData-ABNF].

Dynamic properties can be used in the same way as declared properties. If they are not defined on an instance, they evaluate to null
		 



On resources that derive from directoryObject, $count is only supported in an advanced query. See Advanced query capabilities in Azure AD directory objects.
Use of $count is not supported in Azure AD B2C tenants.


	! Escaping single quotes
For requests that use single quotes, if any parameter values also contain single quotes, those must be double escaped; otherwise, the request will fail due to invalid syntax. In the example, the string value let''s meet for lunch? has the single quote escaped.

		Advanced query capabilities on Azure AD objects https://learn.microsoft.com/en-us/graph/aad-advanced-queries?tabs=http
		 */
	}
}