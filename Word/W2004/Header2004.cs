using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004
{
    public class Header2004 : IHeader
    {
        private const string HEADER_TOP = "\n\n	<w:hdr w:type=\"odd\">";
        private const string HEADER_BOTTON = "\n	</w:hdr>";

        private const string HIDE_HEADER__FOOTER_FIRST_PAGE =
            "\n            <w:hdr w:type=\"first\"> "
            + "\n                <w:p wsp:rsidR=\"00005E72\" wsp:rsidRDefault=\"00005E72\"> "
            + "\n                    <w:pPr> "
            + "\n                        <w:pStyle w:val=\"Header\"/> "
            + "\n                    </w:pPr> "
            + "\n                </w:p> "
            + "\n            </w:hdr> "
            + "\n            <w:ftr w:type=\"first\"> "
            + "\n                <w:p wsp:rsidR=\"00005E72\" wsp:rsidRDefault=\"00005E72\"> "
            + "\n                    <w:pPr> "
            + "\n                        <w:pStyle w:val=\"Footer\"/> "
            + "\n                    </w:pPr> "
            + "\n                </w:p> "
            + "\n            </w:ftr> "
            + "\n            <w:pgSz w:w=\"11900\" w:h=\"16840\"/> "
            +
            "\n            <w:pgMar w:top=\"629\" w:right=\"2517\" w:bottom=\"1259\" w:left=\"1888\" w:header=\"709\" w:footer=\"709\" w:gutter=\"0\"/> "
            + "\n            <w:cols w:space=\"708\"/> "
            + "\n            <w:titlePg/> ";

        #region Implementation of IElement

        /// <summary>
        ///   <p>This method returns the content (XML or HTML) of the Element and the content.</p>
        ///   <p>If you are using W2004, the return will be the XML required to generate the element.</p>
        /// 
        ///   <p>Important: Once you call this method, the Document value is cached an no elements can be added later.</p>
        ///   <example>
        ///     <p>This is the XML that generates a <code>BreakLine</code>:</p>
        ///     <code>
        ///       <w:p wsp:rsidR = '008979E8' wsp:rsidRDefault = '008979E8' />
        ///     </code>
        ///   </example>
        /// </summary>
        /// <value>This is the string value of the element ready to be appended/inserted in the Document.</value>
        public string Content
        {
            get
            {
                if ("".Equals(_txt.ToString()))
                {
                    return "";
                }
                if (_hasBeenCalledBefore)
                {
                    return _txt.ToString();
                }
                else
                {
                    _hasBeenCalledBefore = true;
                }

                _txt.Insert(0, HEADER_TOP);
                _txt.Append(HEADER_BOTTON);

                return _txt.ToString();
            }
        }

        #endregion

        #region Implementation of IHasElement

        /// <summary>
        ///   The content of the Element will be evaluated inside the object.
        /// 
        ///   This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name = "e"></param>
        public void AddEle(IElement e)
        {
            this._txt.Append("\n" + e.Content);
        }

        /// <summary>
        ///   In order to give flexibility, string methods was added.
        ///   Imagine you need to add an element which hasn't been implemented, for example, <b>Graphs</b>.
        ///   You know how to generate the XML to generate a Graph because you figured it out by yourself.
        ///   In this case you could do this:
        ///   <code>
        ///     IDocument myDoc = new Document2004();
        ///     string myXml = myMethodToReturnXML(); //This method return the XML to generate my graph - you have to write this one
        ///     myDoc.getBody().addEle(new Paragraph("This is a my graph: " + myXml));
        ///   </code>
        ///   This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name = "str"></param>
        public void AddEle(string str)
        {
            this._txt.Append("\n" + str);
        }

        #endregion

        #region Implementation of IHeader

        public void SetHideHeaderAndFooterFirstPage(bool value)
        {
            this._hideHeaderAndFooterFirstPage = value;
        }

        public bool GetHideHeaderAndFooterFirstPage()
        {
            return this._hideHeaderAndFooterFirstPage;
        }

        public string GetHideHeaderAndFooterFirstPageXml()
        {
            return HIDE_HEADER__FOOTER_FIRST_PAGE;
        }

        #endregion

        private readonly StringBuilder _txt = new StringBuilder("");

        private bool _hasBeenCalledBefore;
        // if getContent has already been called, I cached the result for future invocations

        private bool _hideHeaderAndFooterFirstPage;
    }
}