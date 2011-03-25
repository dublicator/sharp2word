using System.Collections;
using System.Text;
using Word.Api.Interfaces;
using Word.W2004.Elements.TableElements;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Use this class to create paragraphs. 
    ///   <see cref = "ParagraphPiece" />
    /// </summary>
    public class Paragraph : IElement, IFluentElement<Paragraph>, IFluentElementStylable<ParagraphStyle>
    {
        public ParagraphPiece[] _pieces;

        private readonly ArrayList _tabs = new ArrayList();
        private ParagraphStyle _style = new ParagraphStyle();


        /// <summary>
        /// </summary>
        /// <param name = "value">string for a simple Paragraph. Assuming that you don't want to apply style on part of this text.</param>
        public Paragraph(string value)
        {
            if (value == null || "".Equals(value))
            {
                return;
            }
            ParagraphPiece piece = ParagraphPiece.With(value);
            _pieces = new ParagraphPiece[1];
            _pieces[0] = piece;
        }

        /// <summary>
        ///   It receives many ParagraphPieces with their own style/formating
        /// </summary>
        /// <param name = "pieces"></param>
        public Paragraph(params ParagraphPiece[] pieces)
        {
            this._pieces = pieces;
        }


        public ParagraphStyle Style
        {
            set { this._style = value; }
        }

        #region IElement Members

        public string Content
        {
            get
            {
                if (_pieces == null)
                {
                    // || pieces.length == 0){
                    return "";
                }

                StringBuilder sb = new StringBuilder("");
                foreach (ParagraphPiece piece in _pieces)
                {
                    sb.Append(piece.Content);
                }

                string txt =
                    "	<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"00000000\">"
                    + "\n		{styleText}" // {styleText} is inside styleText
                    + "\n		{tabs}"
                    + "\n		{value}"
                    + "\n	</w:p>";

                if ("".Equals(sb.ToString()))
                {
                    //if there is no content in the pieces, there is no return - just empty string.
                    return "";
                }
                else
                {
                    //For convention, it should be the last thing before returning the xml content.
                    txt = _style.GetNewContentWithStyle(txt);

                    string addTab = "";
                    if (_tabs != null && !(_tabs.Count == 0))
                    {
                        addTab = "  <w:pPr>" + "\n    <w:tabs>";
                        foreach (Tab tab in _tabs)
                        {
                            addTab += "\n        <w:tab w:val=\"" + tab.Align.getValue() + "\" w:pos=\"" + tab.Position +
                                      "\" />";
                        }
                        addTab += "\n    </w:tabs>" + "\n </w:pPr>";
                    }
                    txt = txt.Replace("{tabs}", addTab);


                    return txt.Replace("{value}", sb.ToString());
                }
            }
        }

        #endregion

        #region IFluentElement<Paragraph> Members

        public Paragraph Create()
        {
            return this;
        }

        #endregion

        #region IFluentElementStylable<ParagraphStyle> Members

        public ParagraphStyle WithStyle()
        {
            _style.Element = this;
            return _style;
        }

        #endregion

        /// <summary>
        /// Created a Paragraph with a simple @ParagraphPiece inside
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Paragraph With(string value)
        {
            //if(value == null || "".equals(value)){
            //return null;
            //}
            Paragraph par = new Paragraph();
            ParagraphPiece piece = ParagraphPiece.With(value);
            par._pieces = new ParagraphPiece[1];
            par._pieces[0] = piece;
            return par;
            //return new Paragraph(value);
        }

        public static Paragraph WithPieces(params ParagraphPiece[] pieces)
        {
            return new Paragraph(pieces);
        }

        /// <summary>
        ///   Configures the Align and position of the tabs ('\t'). Position is pretty much the size of EACH tab or each '\t'.
        /// </summary>
        /// <param name = "tabAlign">Right or Left according to the Enum @TabAlign</param>
        /// <param name = "position">Kind of size of EACH tab or each '\t'</param>
        /// <returns>The fluent actual paragraph</returns>
        public Paragraph AddTab(TabAlign tabAlign, int position)
        {
            _tabs.Add(new Tab(tabAlign, position));
            return this;
        }

        #region Nested type: Tab

        private class Tab
        {
            private readonly TabAlign _align;
            private readonly int _position;

            public Tab(TabAlign pAlign, int pPosition)
            {
                _align = pAlign;
                _position = pPosition;
            }

            public TabAlign Align
            {
                get { return _align; }
            }

            public int Position
            {
                get { return _position; }
            }
        }

        #endregion

        #region Nested type: TabAlign

        public class TabAlign
        {
            private readonly string _value;


            private TabAlign(string value)
            {
                this._value = value;
            }

            public static TabAlign LEFT
            {
                get { return new TabAlign("left"); }
            }

            public static TabAlign RIGHT
            {
                get { return new TabAlign("right"); }
            }

            public string getValue()
            {
                return _value;
            }
        }

        #endregion
    }
}