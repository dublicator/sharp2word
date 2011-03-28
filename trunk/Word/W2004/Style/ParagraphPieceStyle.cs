using System.Drawing;
using System.Text;
using Word.Api.Interfaces;
using Word.Utils;
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
        private bool _caps;
        private Color _color;
        private bool _doubleStrike;
        private bool _emboss;
        private Font _font;
        private string _fontSize = "";
        private bool _imprint;
        private bool _italic;
        private bool _outline;
        private bool _shadow;
        private bool _smallCaps;
        private bool _strike;
        private bool _subscript;
        private bool _superscript;
        private string _textColor = "";
        private bool _underline;
        private bool _vanish;

        #region ISuperStylin Members

        public override string GetNewContentWithStyle(string txt)
        {
            StringBuilder style = new StringBuilder("");

            // 'doStyleFont' has to be before 'doStyleBold' and 'doStyleItalic'
            // because of the 'smart bold/italic' based on font type.
            DoStyleFont(style);
            DoStyleBold(style);
            DoStyleItalic(style);
            DoStyleUnderline(style);

            DoStyleSubscript(style);
            DoStyleSuperscript(style);
            DoStyleStrike(style);
            DoStyleCaps(style);
            DoStyleDoubleStrike(style);
            DoStyleEmboss(style);
            DoStyleImprint(style);
            DoStyleOutline(style);
            DoStyleShadow(style);
            DoStyleSmallCaps(style);
            DoStyleVanish(style);

            DoStyleTextColorHexa(style);
            DoStyleColorEnum(style);
            DoStyleFontSize(style);
            DoStyleBgColor(style);

            return DoStyleReplacement(style, txt);
        }

        #endregion

        public ParagraphPieceStyle DoubleStrike()
        {
            _doubleStrike = true;
            return this;
        }

        public ParagraphPieceStyle Shadow()
        {
            _shadow = true;
            return this;
        }

        public ParagraphPieceStyle Outline()
        {
            _outline = true;
            return this;
        }

        public ParagraphPieceStyle Emboss()
        {
            _emboss = true;
            return this;
        }

        public ParagraphPieceStyle Imprint()
        {
            _imprint = true;
            return this;
        }

        /// <summary>
        /// It makes capital case but with middle letters a bit smaller the the first one in the word
        /// </summary>
        /// <returns></returns>
        public ParagraphPieceStyle SmallCaps()
        {
            _smallCaps = true;
            return this;
        }

        /// <summary>
        /// It makes the value Capital case
        /// </summary>
        /// <returns></returns>
        public ParagraphPieceStyle Caps()
        {
            _caps = true;
            return this;
        }

        /// <summary>
        /// It makes the text hidden or doesn't show it.
        /// </summary>
        /// <returns></returns>
        public ParagraphPieceStyle Vanish()
        {
            _vanish = true;
            return this;
        }

        private void DoStyleDoubleStrike(StringBuilder style)
        {
            if (this._doubleStrike)
            {
                style.Append("\n			<w:dstrike/>");
            }
        }

        private void DoStyleShadow(StringBuilder style)
        {
            if (this._shadow)
            {
                style.Append("\n			<w:shadow/>");
            }
        }

        private void DoStyleOutline(StringBuilder style)
        {
            if (this._outline)
            {
                style.Append("\n			<w:outline/>");
            }
        }

        private void DoStyleEmboss(StringBuilder style)
        {
            if (this._emboss)
            {
                style.Append("\n			<w:emboss/>");
            }
        }

        private void DoStyleImprint(StringBuilder style)
        {
            if (this._imprint)
            {
                style.Append("\n			<w:imprint/>");
            }
        }

        private void DoStyleSmallCaps(StringBuilder style)
        {
            if (this._smallCaps)
            {
                style.Append("\n			<w:smallCaps/>");
            }
        }

        private void DoStyleCaps(StringBuilder style)
        {
            if (this._caps)
            {
                style.Append("\n			<w:caps/>");
            }
        }

        private void DoStyleVanish(StringBuilder style)
        {
            if (this._vanish)
            {
                style.Append("\n			<w:vanish/>");
            }
        }

        /// <summary>
        ///   Set the text to Bold
        /// </summary>
        /// <returns></returns>
        public ParagraphPieceStyle Bold()
        {
            _bold = true;
            return this;
        }

        public ParagraphPieceStyle Italic()
        {
            this._italic = true;
            return this;
        }

        public ParagraphPieceStyle Underline()
        {
            this._underline = true;
            return this;
        }

        public ParagraphPieceStyle Subscript()
        {
            _subscript = true;
            return this;
        }

        public ParagraphPieceStyle Superscript()
        {
            _superscript = true;
            return this;
        }

        public ParagraphPieceStyle Strike()
        {
            _strike = true;
            return this;
        }

        private void DoStyleBgColor(StringBuilder style)
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
        public ParagraphPieceStyle BgColor(string bgColor)
        {
            this._bgColor = bgColor;
            return this;
        }

        private void DoStyleBold(StringBuilder style)
        {
            if (this._bold)
            {
                style.Append("\n            	<w:b/>");
            }
        }

        private void DoStyleSubscript(StringBuilder style)
        {
            if (this._subscript)
            {
                style.Append("\n			<w:vertAlign w:val=\"subscript\"/>");
            }
        }

        private void DoStyleSuperscript(StringBuilder style)
        {
            if (this._superscript)
            {
                style.Append("\n			<w:vertAlign w:val=\"superscript\"/>");
            }
        }

        private void DoStyleStrike(StringBuilder style)
        {
            if (_strike)
            {
                style.Append("\n			<w:strike/>");
            }
        }

        private void DoStyleItalic(StringBuilder style)
        {
            if (this._underline)
            {
                style.Append("\n			<w:u w:val=\"single\"/>");
            }
        }

        private void DoStyleUnderline(StringBuilder style)
        {
            if (this._underline)
            {
                style.Append("\n            	<w:i/>");
            }
        }

        private void DoStyleTextColorHexa(StringBuilder style)
        {
            if (!this._textColor.Equals(""))
            {
                style.Append("\n			<w:color w:val=\"" + this._textColor + "\"/>");
            }
        }

        private void DoStyleColorEnum(StringBuilder style)
        {
            var clr = ImageUtils.ColorToHex(this._color);
            if (!string.IsNullOrEmpty(clr) && !clr.Equals(""))
            {
                style.Append("\n			<w:color w:val=\"" + clr + "\"/>");
            }
        }

        private void DoStyleFont(StringBuilder style)
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

        private void DoStyleFontSize(StringBuilder style)
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

        private static string DoStyleReplacement(StringBuilder style, string txt)
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

        public new ParagraphPiece Create()
        {
            return (ParagraphPiece) base.Create();
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
        public ParagraphPieceStyle TextColor(string textColor)
        {
            this._textColor = textColor;
            return this;
        }

        /// <summary>
        ///   Set text color from the Enum @Color, case you don't know any hexa code color
        /// </summary>
        /// <param name = "color"></param>
        /// <returns></returns>
        public ParagraphPieceStyle TextColor(Color color)
        {
            this._color = color;
            return this;
        }

        public ParagraphPieceStyle Font(Font font)
        {
            this._font = font;
            return this;
        }

        /// <summary>
        ///   Specify the font size the same way you see on MW Word.
        /// </summary>
        /// <param name = "size"></param>
        /// <returns></returns>
        public ParagraphPieceStyle FontSize(int size)
        {
            this._fontSize = (size*2).ToString();
            return this;
        }
    }
}