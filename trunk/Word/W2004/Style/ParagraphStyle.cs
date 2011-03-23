using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004.Style
{
    /// <summary>
    ///   Use this class to apply style to the Paragraph - the whole paragraph. At the moment is only possible to apply Align. 
    ///   If you want to apply bold, italic or underline, use a ParagraphPiece to do that.
    /// </summary>
    public class ParagraphStyle : AbstractStyle, ISuperStylin
    {
        private Align _align = Style.Align.LEFT;
        private string _bgColor = "";

        #region ISuperStylin Members

        public override string GetNewContentWithStyle(string txt)
        {
            StringBuilder style = new StringBuilder("");

            //There will be always align 'Left' by default
            DoStyleAlignment(style);
            DoStyleBgColor(style);

            return DoStyleReplacement(style, txt);
        }

        #endregion

        private void DoStyleAlignment(StringBuilder style)
        {
            style.Append("  <w:jc w:val=\"" + _align.Value + "\"/> \n    " + "       {styleText}\n   ");
        }

        private void DoStyleBgColor(StringBuilder style)
        {
            if (!_bgColor.Equals(""))
            {
                style.Append("\n            <w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"" + _bgColor + "\" />\n");
            }
        }

        private static string DoStyleReplacement(StringBuilder style, string txt)
        {
            // if (!"".equals(style.toString())) {
            style.Insert(0, "\n  <w:pPr>");
            style.Append("\n     </w:pPr>");
            txt = txt.Replace("{styleText}", style.ToString()); // Convention:
            // apply styles
            // }

            //txt = txt.replace("{styleText}", styleValue);
            txt = txt.Replace("[{]style(.*)[}]", ""); //Convention: replace unused styles after...
            return txt;
        }

        /// <summary>
        ///   If you know the color code, just to straight to the point! Eg.: yellow:
        ///   FFFF00, black: 000000, red: FF0000, blue: 0000FF, green: 008000, etc...
        /// 
        ///   If you want, you can use the class Color.whatever_color.
        /// </summary>
        /// <param name = "bgColor">Hexadecimal color code</param>
        /// <returns></returns>
        public ParagraphStyle BgColor(string bgColor)
        {
            this._bgColor = bgColor;
            return this;
        }

        /// <summary>
        /// Sets align for the whole paragraph
        /// </summary>
        /// <param name="align"></param>
        /// <returns></returns>
        public ParagraphStyle Align(Align align)
        {
            this._align = align;
            return this;
        }
    }
}