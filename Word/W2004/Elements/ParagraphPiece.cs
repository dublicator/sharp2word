using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Use this class ONLY inside Paragraph in order to format pieces of a paragraph.
    ///   for example, if you want to make one and only one word of the paragraph bold.
    ///   ALWAYS USE THIS CLASS INSIDE A PARAGRAPH. OTHERWISE WON'T WORK. This is the way you should use it:
    ///   <code>
    ///     myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with("This is one piece.").create()));
    ///   </code>
    ///   <code>
    ///     myDoc.addEle(ParagraphPiece.with("This is one piece.").create());
    ///   </code>
    /// </summary>
    public class ParagraphPiece : IElement, IFluentElement<ParagraphPiece>, IFluentElementStylable<ParagraphPieceStyle>
    {
        private string _value = "";
        private ParagraphPieceStyle _style = new ParagraphPieceStyle();

        private string _txt =
            "\n		<w:r>"
            + "\n			{styleText}"
            + "\n			<w:t>{value}</w:t>"
            + "\n		</w:r>";

        public ParagraphPiece(string value)
        {
            this._value = value;
        }

        public ParagraphPiece()
        {
            
        }

        public ParagraphPieceStyle Style
        {
            get { return _style; }
            set { this._style = value; }
        }

        #region IElement Members

        public string Content
        {
            get
            {
                if ("".Equals(this._value) || this._value == null)
                {
                    // null is very unusual. That the reason null comparison is after empty verification. I am not sure if we use ApacheUtils we can achieve the same  
                    return "";
                }

                //For convention, it should be the last thing before returning the xml content.
                _txt = _style.GetNewContentWithStyle(_txt);

                return _txt.Replace("{value}", this._value);
            }
        }

        #endregion

        #region IFluentElement<ParagraphPiece> Members

        public ParagraphPiece Create()
        {
            return this;
        }

        #endregion

        #region IFluentElementStylable<ParagraphPieceStyle> Members

        public ParagraphPieceStyle WithStyle()
        {
            this._style.Element = this;
            return this._style;
        }

        #endregion

        public static ParagraphPiece With(string value)
        {
            ParagraphPiece par = new ParagraphPiece();
            par._value = value;
            return par;
            //return new ParagraphPiece(value);
        }
    }
}