using Word.Api.Interfaces;

namespace Word.W2004.Style
{
    /// <summary>
    /// Use this class to apply style to the Paragraph - the whole paragraph. At the moment is only possible to apply Align. 
    /// If you want to apply bold, italic or underline, use a ParagraphPiece to do that. 
    /// </summary>
    public class ParagraphStyle : AbstractStyle, ISuperStylin
    {
        private Align align = Align.LEFT;

        public override string getNewContentWithStyle(string txt)
        {
            string styleValue = "\n		<w:pPr>\n		" +
                                "	<w:jc w:val=\"" + align.getValue() + "\"/> \n	" +
                                "		{styleText}\n 	" +
                                "	</w:pPr>";
            txt = txt.Replace("{styleAlign}", styleValue);
            txt = txt.Replace("[{]style(.*)[}]", ""); //Convention: replace unused styles after... 
            return txt;
        }

        public ParagraphStyle setAlign(Align align)
        {
            this.align = align;
            return this;
        }
    }
}