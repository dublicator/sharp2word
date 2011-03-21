using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    /// Use this class ONLY inside Paragraph in order to format pieces of a paragraph.
    /// for example, if you want to make one and only one word of the paragraph bold.
    /// ALWAYS USE THIS CLASS INSIDE A PARAGRAPH. OTHERWISE WON'T WORK. This is the way you should use it:
    /// <code>
    ///     myDoc.addEle(Paragraph.withPieces(ParagraphPiece.with("This is one piece.").create()));
    /// </code>
    /// <code>
    ///     myDoc.addEle(ParagraphPiece.with("This is one piece.").create());
    /// </code>
    /// </summary>
    public class ParagraphPiece : IElement, IFluentElement<ParagraphPiece>, IFluentElementStylable<ParagraphPieceStyle>
    {
        private string value = "";
        private ParagraphPieceStyle style = new ParagraphPieceStyle();

        string txt =
         "\n		<w:r>"
        + "\n			{styleText}"
        + "\n			<w:t>{value}</w:t>"
        + "\n		</w:r>";

        public ParagraphPiece(string value)
        {
            this.value = value;
        }

        public string getContent()
        {
            if ("".Equals(this.value) || this.value == null)
            { // null is very unusual. That the reason null comparison is after empty verification. I am not sure if we use ApacheUtils we can achieve the same  
                return "";
            }

            //For convention, it should be the last thing before returning the xml content.
            txt = style.getNewContentWithStyle(txt);

            return txt.Replace("{value}", this.value);
        }

        //### Gettets and Setters
        public ParagraphPieceStyle getStyle()
        {
            return style;
        }
        public void setStyle(ParagraphPieceStyle style)
        {
            this.style = style;
        }

        public ParagraphPieceStyle withStyle()
        {
            this.style.setElement(this);
            return this.style;
        }

        public static ParagraphPiece with(string value)
        {
            return new ParagraphPiece(value);
        }

        public ParagraphPiece create()
        {
            return this;
        }
    }
}