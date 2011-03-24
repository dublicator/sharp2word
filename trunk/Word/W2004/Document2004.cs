using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using Word.Api.Interfaces;
using Word.Utils;

namespace Word.W2004
{
    public class Document2004 : IDocument, IElement
    {
        private static bool _validated = true;
        private readonly IBody _body = new Body2004();
        private readonly IHead _head = new Head2004();
        private readonly StringBuilder _txt = new StringBuilder();

        private bool _hasBeenCalledBefore;
        // if getContent has already been called, I cached the result for future invocations

        private bool _isLandscape;

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
                if (_hasBeenCalledBefore)
                {
                    return _txt.ToString();
                }
                else
                {
                    _hasBeenCalledBefore = true;
                }
                _txt.Append(this.Uri);
                _txt.Append(this.Head.Content);

                _txt.Append(this.Body.Content);

                _txt.Append("\n</w:wordDocument>");

                string finalString = SetUpPageOrientation(_txt.ToString());

                return finalString;
            }
        }


        public void SetPageOrientationLandscape()
        {
            this._isLandscape = true;
        }

        /// <summary>
        /// Save document to local disk
        /// </summary>
        /// <param name="path">Path to file</param>
        public void Save(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create))
            {
                using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                {
                    w.Write(Util.Pretty(Content));
                }
            }

        }

        public IBody Body
        {
            get { return this._body; }
        }

        public IHead Head
        {
            get { return this._head; }
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
        public void AddEle(IElement e)
        {
            this.Body.AddEle(e);
        }

        public void AddEle(string str)
        {
            this.Body.AddEle(str);
        }

        #endregion

        private string SetUpPageOrientation(string txt)
        {
            if (this._isLandscape)
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

        public bool Validate()
        {
            // Set the validation settings.
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.Schema;
            settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
            settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            settings.ValidationEventHandler += ValidationCallBack;


            // Create the XmlReader object.
            XmlReader reader = XmlReader.Create("inlineSchema.xml", settings);

            // Parse the file. 
            while (reader.Read() || !_validated) ;
            return _validated;
        }

        // Display any warnings or errors.
        private static void ValidationCallBack(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
            {
            }
            else
            {
                _validated = false;
            }
        }
    }
}