using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004
{
    public class Body2004 : IBody
    {
        private readonly IFooter _footer = new Footer2004();
        private readonly IHeader _header = new Header2004();
        private readonly StringBuilder _txt = new StringBuilder("");

        private const string DEFAULT_MGR_TOP = "629";
        private const string DEFAULT_MGR_BOTTOM = "1259";
        private const string DEFAULT_MGR_RIGHT = "2517";
        private const string DEFAULT_MGR_LEFT = "1888";
        private int _pageMgrLeft = -1;
        private int _pageMgrRight = -1;
        private int _pageMgrTop = -1;
        private int _pageMgrBottom = -1;
        private bool _pageMarginCustomSetting;

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
                StringBuilder res = new StringBuilder();
                res.Append("\n<w:body>");

                res.Append(_txt.ToString());

                string header = this.Header.Content;
                string footer = this.Footer.Content;
                if (!"".Equals(header) || !"".Equals(footer) || _pageMarginCustomSetting)
                {
                    const string headerFooterTop = "<w:sectPr wsp:rsidR=\"00DB1FE5\" wsp:rsidSect=\"00471A86\">";
                    const string headerFooterBotton = "</w:sectPr>";

                    res.Append("\n" + headerFooterTop);
                    res.Append(header); //header has to be inside the w:body
                    res.Append(footer); //header has to be inside the w:body
                    if (this.Header.GetHideHeaderAndFooterFirstPage())
                    {
                        res.Append(this.Header.GetHideHeaderAndFooterFirstPageXml());
                    }
                    string margin = PAGE_MARGIN;
                    margin = margin.Replace("{mgr_top}", _pageMgrLeft == -1 ? DEFAULT_MGR_TOP : _pageMgrTop.ToString());
                    margin = margin.Replace("{mgr_bottom}", _pageMgrBottom == -1 ? DEFAULT_MGR_BOTTOM : _pageMgrTop.ToString());
                    margin = margin.Replace("{mgr_right}", _pageMgrRight == -1 ? DEFAULT_MGR_RIGHT : _pageMgrTop.ToString());
                    margin = margin.Replace("{mgr_left}", _pageMgrLeft == -1 ? DEFAULT_MGR_LEFT : _pageMgrTop.ToString());
                    res.Append(margin);//page amrgin setting    

                    res.Append("\n" + headerFooterBotton);
                }

                res.Append("\n</w:body>");
                return res.ToString();
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

        #region Implementation of IBody

        public IHeader Header
        {
            get { return _header; }
        }

        public IFooter Footer
        {
            get { return _footer; }
        }

        public void SetMarginBody(double top, double bottom, double left, double right)
        {
            _pageMgrLeft = (int) (56.7*left);
            _pageMgrRight = (int) (56.7 * right);
            _pageMgrTop = (int) (56.7 * top);
            _pageMgrBottom = (int) (56.7 * bottom);
            _pageMarginCustomSetting = true;
        }

        #endregion

        private const string PAGE_MARGIN =
            "		<w:pgMar w:top=\"{mgr_top}\" "
            + "\n		w:right=\"{mgr_right}\" "
            + "\n		w:bottom=\"{mgr_bottom}\" "
            + "\n		w:left=\"{mgr_left}\" "
            + "\n		w:header=\"709\" w:footer=\"709\" w:gutter=\"0\"/> ";
    }
}