
using System.Collections;
using System.Text;
using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Use this class to create paragraphs. 
    ///   <see cref = "ParagraphPiece" />
    /// </summary>
    public class Paragraph : IElement, IFluentElement<Paragraph>, IFluentElementStylable<ParagraphStyle>
    {
        private readonly ParagraphPiece[] pieces;

        private ParagraphStyle _style = new ParagraphStyle();

        private ArrayList tabs = new ArrayList();


        /// <summary>
        /// </summary>
        /// <param name = "value">string for a simple Paragraph. Assuming that you don't want to apply style on part of this text.</param>
        public Paragraph(string value)
        {
            if (value == null || "".Equals(value))
            {
                return;
            }
            ParagraphPiece piece = new ParagraphPiece(value);
            pieces = new ParagraphPiece[1];
            pieces[0] = piece;
        }

        /// <summary>
        ///   It receives many ParagraphPieces with their own style/formating
        /// </summary>
        /// <param name = "pieces"></param>
        public Paragraph(params ParagraphPiece[] pieces)
        {
            this.pieces = pieces;
        }


        //## Getters and Setters

        public ParagraphStyle Style
        {
            get { return _style; }
        }

        #region IElement Members

        public string Content
        {
            get
            {
                if (pieces == null)
                {
                    // || pieces.length == 0){
                    return "";
                }

                StringBuilder sb = new StringBuilder("");
                foreach (ParagraphPiece piece in pieces)
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
                    txt = _style.getNewContentWithStyle(txt);

                    string addTab = "";
                    if (tabs != null && !(tabs.Count == 0))
                    {
                        addTab = "  <w:pPr>" + "\n    <w:tabs>";
                        foreach (Tab tab in tabs)
                        {
                            addTab += "\n        <w:tab w:val=\"" + tab.getAlign().getValue() + "\" w:pos=\"" + tab.getPosition() + "\" />";
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

        public Paragraph create()
        {
            return this;
        }

        #endregion

        #region IFluentElementStylable<ParagraphStyle> Members

        public ParagraphStyle withStyle()
        {
            _style.setElement(this);
            return _style;
        }

        #endregion

        public void setStyle(ParagraphStyle style)
        {
            this._style = style;
        }

        public static Paragraph with(string value)
        {
            return new Paragraph(value);
        }

        public static Paragraph withPieces(params ParagraphPiece[] pieces)
        {
            return new Paragraph(pieces);
        }

        /// <summary>
        /// Configures the Align and position of the tabs ('\t'). Position is pretty much the size of EACH tab or each '\t'.
        /// </summary>
        /// <param name="tabAlign">Right or Left according to the Enum @TabAlign</param>
        /// <param name="position">Kind of size of EACH tab or each '\t'</param>
        /// <returns>The fluent actual paragraph</returns>
        public Paragraph addTab(TabAlign tabAlign, int position)
        {
            tabs.Add(new Tab(tabAlign, position));
            return this;
        }

        public class TabAlign
        {
            public static TabAlign LEFT
            {
                get { return new TabAlign("left"); }
            }

            public static TabAlign RIGHT
            {
                get { return new TabAlign("right"); }
            }

            private string value;


            private TabAlign(string value)
            {
                this.value = value;
            }

            public string getValue()
            {
                return value;
            }
        }

        private class Tab
        {
            private TabAlign align;
            private int position;

            public Tab(TabAlign pAlign, int pPosition)
            {
                align = pAlign;
                position = pPosition;
            }

            public TabAlign getAlign()
            {
                return align;
            }

            public int getPosition()
            {
                return position;
            }
        }
    }
}