using Microsoft.Graph;
using Sharepoint.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dotnet_console_microsoft_graph
{
    public class SharePointIntegrationTestsIO
    {
        private SharePointUtl_V5 _sharepointutility;
        public SharePointIntegrationTestsIO(GraphServiceClient graphServiceClient)
        {
            _sharepointutility = new(graphServiceClient);
        }

        public bool RunAllTests() {
            return false;
        }
        public bool PrepareSharepointForTesting()
        {
            return false;
        }
    }
}
