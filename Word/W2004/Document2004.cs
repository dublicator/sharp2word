using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004
{
    public class Document2004 : IDocument, IElement
    {
        private readonly IBody body = new Body2004();
        private readonly IHead head = new Head2004();
        private readonly StringBuilder _txt = new StringBuilder();

        private bool hasBeenCalledBefore;
                     // if getContent has already been called, I cached the result for future invocations

        private bool isLandscape;

        #region IDocument Members

        public string Uri
        {
            get
            {
                const string uri = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?> "
                                   + "<?mso-application progid=\"Word.Document\"?> "
                                   + "<w:wordDocument xmlns:aml=\"http://schemas.microsoft.com/aml/2001/core\" "
                                   +
                                   " xmlns:dt=\"uuid:C2F41010-65B3-11d1-A29F-00AA00C14882\" xmlns:mo=\"http://schemas.microsoft.com/office/mac/office/2008/main\" "
                                   + " xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                                   +
                                   " xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
                                   +
                                   " xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
                                   + " xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\" "
                                   + " xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\" "
                                   + " xmlns:wsp=\"http://schemas.microsoft.com/office/word/2003/wordml/sp2\" "
                                   + " xmlns:sl=\"http://schemas.microsoft.com/schemaLibrary/2003/core\" "
                                   + " w:macrosPresent=\"no\" w:embeddedObjPresent=\"no\" w:ocxPresent=\"no\" "
                                   + " xml:space=\"preserve\"> "
                                   +
                                   " <w:ignoreSubtree w:val=\"http://schemas.microsoft.com/office/word/2003/wordml/sp2\" /> ";
                return uri;
            }
        }

        public string Content
        {
            get
            {
                if (hasBeenCalledBefore)
                {
                    return _txt.ToString();
                }
                else
                {
                    hasBeenCalledBefore = true;
                }
                _txt.Append(this.Uri);
                _txt.Append(this.Head.Content);

                _txt.Append(this.Body.Content);

                _txt.Append("\n</w:wordDocument>");

                string finalString = setUpPageOrientation(_txt.ToString());

                return finalString;
            }
        }


        public void setPageOrientationLandscape()
        {
            this.isLandscape = true;
        }

        public IBody Body
        {
            get { return this.body; }
        }

        public IHead Head
        {
            get { return this.head; }
        }

        public IFooter Footer
        {
            get
            {
                //forward it to the body
                return this.Body.Footer;
            }
        }

        public IHeader Header
        {
            get { return this.Body.Header; //forward it to the body
            }
        }

        /// <summary>
        ///   This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name = "e"></param>
        public void addEle(IElement e)
        {
            this.Body.addEle(e);
        }

        public void addEle(string str)
        {
            this.Body.addEle(str);
        }

        #endregion

        private string setUpPageOrientation(string txt)
        {
            if (this.isLandscape)
            {
                const string orientation = "    <w:sectPr wsp:rsidR=\"00F04FB2\" wsp:rsidSect=\"00146B2A\">\n"
                                           + "      <w:pgSz w:w=\"16834\" w:h=\"11904\" w:orient=\"landscape\"/>\n"
                                           +
                                           "      <w:pgMar w:top=\"1800\" w:right=\"1440\" w:bottom=\"1800\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>\n"
                                           + "      <w:cols w:space=\"708\"/>\n" + "    </w:sectPr>";
                txt = txt.Replace("</w:body>", orientation + "\n</w:body>");
            }
            return txt;
        }

        public override string ToString()
        {
            return this.Content;
        }
    }
}