
using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004
{
    public class Body2004 : IBody
    {
        StringBuilder txt = new StringBuilder("");
        IHeader header = new Header2004();
        IFooter footer = new Footer2004();
	
        #region Implementation of IElement

        /// <summary>
        /// <p>This method returns the content (XML or HTML) of the Element and the content.</p>
        /// <p>If you are using W2004, the return will be the XML required to generate the element.</p>
        /// 
        /// <p>Important: Once you call this method, the Document value is cached an no elements can be added later.</p>
        /// 
        /// <example>
        ///     <p>This is the XML that generates a <code>BreakLine</code>:</p>
        ///     <code>
        ///         <w:p wsp:rsidR='008979E8' wsp:rsidRDefault='008979E8'/>
        ///     </code>
        /// </example>
        /// </summary>
        /// <returns>This is the String value of the element ready to be appended/inserted in the Document.</returns>
        public string getContent()
        {
            StringBuilder res = new StringBuilder();
            res.Append("\n<w:body>");

            res.Append(txt.ToString());

            string header = this.getHeader().getContent();
            string footer = this.getFooter().getContent();
            if (!"".Equals(header) || !"".Equals(footer))
            {
                string header_footer_top = "<w:sectPr wsp:rsidR=\"00DB1FE5\" wsp:rsidSect=\"00471A86\">";
                string header_footer_botton = "</w:sectPr>";

                res.Append("\n" + header_footer_top);
                res.Append(header);//header has to be inside the w:body
                res.Append(footer);//header has to be inside the w:body
                if (this.getHeader().getHideHeaderAndFooterFirstPage())
                {
                    res.Append(this.getHeader().getHideHeaderAndFooterFirstPageXml());
                }

                res.Append("\n" + header_footer_botton);
            }

            res.Append("\n</w:body>");
            return res.ToString();
        }

        #endregion

        #region Implementation of IHasElement

        /// <summary>
        /// The content of the Element will be evaluated inside the object.
        /// 
        /// This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name="e"></param>
        public void addEle(IElement e)
        {
            this.txt.Append("\n" + e.getContent()); 
        }

        /// <summary>
        /// In order to give flexibility, String methods was added.
        /// Imagine you need to add an element which hasn't been implemented, for example, <b>Graphs</b>.
        /// You know how to generate the XML to generate a Graph because you figured it out by yourself.
        /// In this case you could do this:
        /// <code>
        ///     IDocument myDoc = new Document2004();
        ///     string myXml = myMethodToReturnXML(); //This method return the XML to generate my graph - you have to write this one
        ///     myDoc.getBody().addEle(new Paragraph("This is a my graph: " + myXml));
        /// </code>
        /// This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name="str"></param>
        public void addEle(string str)
        {
            this.txt.Append("\n" + str);
        }

        #endregion

        #region Implementation of IBody

        public IHeader getHeader()
        {
            return header;
        }

        public IFooter getFooter()
        {
            return footer;
        }

        #endregion
    }
}