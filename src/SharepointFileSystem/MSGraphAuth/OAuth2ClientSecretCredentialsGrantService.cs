
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.IdentityModel.Tokens;

namespace MSGraphAuth {
    public class OAuth2ClientSecretCredentialsGrantService {
        private readonly IEnumerable<string> scopes;
        private readonly string clientid;
        private readonly string clientsecret;
        private readonly string tenantid;
        /// <summary>
        /// Constructs a new <see cref="OAuth2ClientSecretCredentialsGrantService"/>.
        /// With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
        /// application permissions need to be set statically (in the portal or by PowerShell), and then granted by
        /// a tenant administrator. 
        ///		Message	"AADSTS1002012: The provided value for scope is not valid. 
        ///		Client credential flows must have a scope value with /.default suffixed to the resource identifier (application ID URI).
        /// </summary>
        /// <param name="clientid">MS Graph application clientid uid</param>
        /// <param name="clientsecret">MS Graph application clientsecret</param>
        /// <param name="tenantid">MS Graph application tenantid</param>
        /// <param name="scopes">List of scopes for the authentication context, defaults to 'apiUrl.default'</param>
        public OAuth2ClientSecretCredentialsGrantService(string clientid, string clientsecret, string tenantid, string? apiUrl, IEnumerable<string>? scopes = null)
        {
            this.scopes = scopes?.ToArray() ?? new string[] { $"{apiUrl}.default" };
            if (string.IsNullOrEmpty(clientid)) throw new ArgumentNullException(nameof(clientid));
            if (string.IsNullOrEmpty(tenantid)) throw new ArgumentNullException(nameof(tenantid));
            if (string.IsNullOrEmpty(clientsecret)) throw new ArgumentNullException(nameof(clientsecret));
            this.clientid = clientid;
            this.clientsecret = clientsecret;
            this.tenantid = tenantid;
        }
        public GraphServiceClient GetClientSecretClient() {

            var options = new TokenCredentialOptions {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                this.tenantid, this.clientid, this.clientsecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            return graphClient;
        }
    }
}

