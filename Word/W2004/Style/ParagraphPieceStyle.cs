using System.Text;
using Word.Api.Interfaces;
using Word.W2004.Elements;

namespace Word.W2004.Style
{
    /// <summary>
    ///   Use this class in order to apply specifics style to paragraph. Eg.:
    ///   one word in bold, other in italic.
    /// </summary>
    public class ParagraphPieceStyle : AbstractStyle, ISuperStylin
    {
        private string _bgColor = "";
        private bool _bold;
        private Color _color;
        private Font _font;
        private string _fontSize = "";
        private bool _italic;
        private string _textColor = "";
        private bool _underline;

        #region ISuperStylin Members

        public override string getNewContentWithStyle(string txt)
        {
            StringBuilder style = new StringBuilder("");

            // 'doStyleFont' has to be before 'doStyleBold' and 'doStyleItalic'
            // because of the 'smart bold/italic' based on font type.
            doStyleFont(style);
            doStyleBold(style);
            doStyleItalic(style);
            doStyleUnderline(style);
            doStyleTextColorHexa(style);
            doStyleColorEnum(style);
            doStyleFontSize(style);
            doStyleBgColor(style);

            return doStyleReplacement(style, txt);
        }

        #endregion

        private void doStyleBgColor(StringBuilder style)
        {
            if (!_bgColor.Equals(""))
            {
                style.Append("\n            <w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"" + _bgColor + "\" />\n");
            }
        }

        /// <summary>
        ///   If you know the color code, just to straight to the point! Eg.: yellow:
        ///   FFFF00, black: 000000, red: FF0000, blue: 0000FF, green: 008000, etc...
        /// 
        ///   If you want, you can use the class Color.whatever_color.
        /// 
        ///   Hexadecimal color code
        /// </summary>
        /// <param name = "bgColor"></param>
        /// <returns></returns>
        public ParagraphPieceStyle setBgColor(string bgColor)
        {
            this._bgColor = bgColor;
            return this;
        }

        private void doStyleBold(StringBuilder style)
        {
            if (this._bold)
            {
                style.Append("\n            	<w:b/>");
            }
        }

        private void doStyleItalic(StringBuilder style)
        {
            if (this._italic)
            {
                style.Append("\n            	<w:i/>");
            }
        }

        private void doStyleUnderline(StringBuilder style)
        {
            if (this._underline)
            {
                style.Append("\n			<w:u w:val=\"single\"/>");
            }
        }

        private void doStyleTextColorHexa(StringBuilder style)
        {
            if (!this._textColor.Equals(""))
            {
                style.Append("\n			<w:color w:val=\"" + this._textColor + "\"/>");
            }
        }

        private void doStyleColorEnum(StringBuilder style)
        {
            if (this._color != null && !this._color.Value.Equals(""))
            {
                style.Append("\n			<w:color w:val=\"" + _color.Value + "\"/>");
            }
        }

        private void doStyleFont(StringBuilder style)
        {
            // Smart Italic/Bold: This will make the font bold/italic according to
            // this.font
            string fontName = "";
            if (this._font != null)
            {
                fontName = this._font.Value;
                if (fontName.Contains("Bold"))
                {
                    this._bold = true;
                }
                else
                {
                    //if is manually 'bold', I also change the font name
                    if (this._bold)
                    {
                        fontName += " Bold";
                    }
                }

                if (fontName.Contains("Italic"))
                {
                    this._italic = true;
                }
                else
                {
                    if (this._italic)
                    {
                        fontName += " Italic";
                    }
                }
            }

            if (this._font != null)
            {
                style.Append("\n			<w:rFonts w:ascii=\"" + fontName + "\" w:h-ansi=\"" + fontName + "\"/>\n");
                style.Append("\n			<wx:font wx:val=\"" + fontName + "\"/>");
            }
        }

        private void doStyleFontSize(StringBuilder style)
        {
            if (!"".Equals(this._fontSize))
            {
                string ffsize = "\n               <w:sz w:val=\"" + this._fontSize
                                + "\" />\n";
                ffsize += "\n               <w:sz-cs w:val=\"" + this._fontSize
                          + "\" />\n";
                style.Append(ffsize);
            }
        }

        private static string doStyleReplacement(StringBuilder style, string txt)
        {
            if (!"".Equals(style.ToString()))
            {
                style.Insert(0, "\n	 <w:rPr>");
                style.Append("\n	 </w:rPr>");
                txt = txt.Replace("{styleText}", style.ToString()); // Convention:
                // apply styles
            }
            // Convention: replace unused styles after...
            txt = txt.Replace("[{]style(.*)[}]", "");
            return txt;
        }

        /**
         * 
         * This is the ParagraphPiece! I am using Covariant Return Type!!! 
         * to be honest, I have never thought how to use and finally here we go!!!
         * It will give the chance to eliminate the necessity of type cast for elements.
         * 
         */

        public new ParagraphPiece create()
        {
            return (ParagraphPiece) base.create();
        }

        public ParagraphPieceStyle setBold(bool bold)
        {
            this._bold = bold;
            return this;
        }

        public ParagraphPieceStyle setItalic(bool italic)
        {
            this._italic = italic;
            return this;
        }

        public ParagraphPieceStyle setUnderline(bool underline)
        {
            this._underline = underline;
            return this;
        }

        /// <summary>
        ///   If you know the color code, just to straight to the point! Eg.:
        ///   <example>
        ///     yellow: FFFF00, black: 000000, red: FF0000, blue: 0000FF, green: 008000, etc...
        ///   </example>
        ///   If you want, you can use the class Color.whatever_color.
        /// </summary>
        /// <param name = "textColor">Hexadecimal color code</param>
        /// <returns></returns>
        public ParagraphPieceStyle setTextColor(string textColor)
        {
            this._textColor = textColor;
            return this;
        }

        public ParagraphPieceStyle setTextColor(Color color)
        {
            this._color = color;
            return this;
        }

        public ParagraphPieceStyle setFont(Font font)
        {
            this._font = font;
            return this;
        }

        /***
        * Pass '50' for something quite big. I am not sure if the unit is pixels or what... 
        * find out later. sorry about that. 
        */

        /// <summary>
        ///   Pass '50' for something quite big. I am not sure if the unit is pixels or what... 
        ///   find out later. sorry about that.
        /// </summary>
        /// <param name = "fontSize"></param>
        /// <returns></returns>
        public ParagraphPieceStyle setFontSize(string fontSize)
        {
            this._fontSize = fontSize;
            return this;
        }
    }
}