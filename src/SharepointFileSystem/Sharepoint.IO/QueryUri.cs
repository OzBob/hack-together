using System.Collections.Specialized;
using System.Web;

namespace Sharepoint.IO
{
    public class QueryUri : Uri
    {
        public QueryUri(string uriString)
            : base(uriString)
        {
        }

        public string Find(string key)
        {
            string queryString = Query;

            if (string.IsNullOrEmpty(queryString)) return string.Empty;

            NameValueCollection queryParams = HttpUtility.ParseQueryString(queryString);

            return queryParams[key] ?? string.Empty;
        }

        public void Add(string key, string value)
        {
            var _ = Find(key);
            if (!string.IsNullOrEmpty(_))
                return;

            string queryString = Query;
            if (string.IsNullOrEmpty(queryString))
                queryString = "?";
            else if (!queryString.StartsWith("?"))
                queryString = "?" + queryString;
            NameValueCollection queryParams = HttpUtility.ParseQueryString(queryString);
            queryParams.Add(key, value);
            string updatedQueryString = queryParams.ToString();
            if (updatedQueryString.StartsWith("?"))
                updatedQueryString = updatedQueryString.Substring(1);
            UriBuilder uriBuilder = new UriBuilder(this)
            {
                Query = updatedQueryString
            };
            //base.(uriBuilder.Uri);
        }
    }
}
