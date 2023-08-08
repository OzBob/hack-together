namespace Sharepoint.IO
{
    public class SPfileDocument
    {
        private string? _trimUri;

        
        public string? SpDeepLinkUrl { get { return _trimUri; } set { _trimUri = value; } }

        public SPfileDocument()
        {
        }

        public SPfileDocument(string trimUri)
        {
            _trimUri = trimUri;
            
        }

        public SPfileDocument(Uri objRec)
        {
            _trimUri = objRec.AbsolutePath;
        }
    }
}

