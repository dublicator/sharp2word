using System.Text;
using Word.Api.Interfaces;
using Word.W2004.Elements;

namespace Word.W2004.Style
{
    /// <summary>
    /// Use this class in order to apply specifics style to paragraph. Eg.:
    /// one word in bold, other in italic.    
    /// </summary>
    public class ParagraphPieceStyle : AbstractStyle, ISuperStylin
    {
        private bool bold = false;
        private bool italic = false;
        private bool underline = false;
        private string textColor = "";
        private Color color;
        public Font font;
        private string fontSize = "";
        private string bgColor = "";

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

        private void doStyleBgColor(StringBuilder style)
        {
            if (!bgColor.Equals(""))
            {
                style.Append("\n            <w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"" + bgColor + "\" />\n");
            }
        }

        /// <summary>
        /// If you know the color code, just to straight to the point! Eg.: yellow:
        /// FFFF00, black: 000000, red: FF0000, blue: 0000FF, green: 008000, etc...
        /// 
        /// If you want, you can use the class Color.whatever_color.
        /// 
        /// Hexadecimal color code
        /// </summary>
        /// <param name="bgColor"></param>
        /// <returns></returns>
        public ParagraphPieceStyle setBgColor(string bgColor)
        {
            this.bgColor = bgColor;
            return this;
        }

        private void doStyleBold(StringBuilder style)
        {
            if (this.bold)
            {
                style.Append("\n            	<w:b/>");
            }
        }

        private void doStyleItalic(StringBuilder style)
        {
            if (this.italic)
            {
                style.Append("\n            	<w:i/>");
            }
        }

        private void doStyleUnderline(StringBuilder style)
        {
            if (this.underline)
            {
                style.Append("\n			<w:u w:val=\"single\"/>");
            }
        }

        private void doStyleTextColorHexa(StringBuilder style)
        {
            if (!this.textColor.Equals(""))
            {
                style.Append("\n			<w:color w:val=\"" + this.textColor + "\"/>");
            }
        }

        private void doStyleColorEnum(StringBuilder style)
        {
            if (this.color != null && !this.color.getValue().Equals(""))
            {
                style.Append("\n			<w:color w:val=\"" + color.getValue() + "\"/>");
            }
        }

        private void doStyleFont(StringBuilder style)
        {
            // Smart Italic/Bold: This will make the font bold/italic according to
            // this.font
            string fontName = "";
            if (this.font != null)
            {
                fontName = this.font.getValue();
                if (fontName.Contains("Bold"))
                {
                    this.bold = true;
                }
                else
                {
                    //if is manually 'bold', I also change the font name
                    if (this.bold)
                    {
                        fontName += " Bold";
                    }
                }

                if (fontName.Contains("Italic"))
                {
                    this.italic = true;
                }
                else
                {
                    if (this.italic)
                    {
                        fontName += " Italic";
                    }
                }
            }

            if (this.font != null)
            {
                style.Append("\n			<w:rFonts w:ascii=\"" + fontName + "\" w:h-ansi=\"" + fontName + "\"/>\n");
                style.Append("\n			<wx:font wx:val=\"" + fontName + "\"/>");
            }
        }

        private void doStyleFontSize(StringBuilder style)
        {
            if (!"".Equals(this.fontSize))
            {
                string ffsize = "\n               <w:sz w:val=\"" + this.fontSize
                        + "\" />\n";
                ffsize += "\n               <w:sz-cs w:val=\"" + this.fontSize
                        + "\" />\n";
                style.Append(ffsize);
            }
        }

        private string doStyleReplacement(StringBuilder style, string txt)
        {
            if (!"".Equals(style.ToString()))
            {
                style.Insert(0, "\n	 <w:rPr>");
                style.Append("\n	 </w:rPr>");
                txt = txt.Replace("{styleText}", style.ToString());// Convention:
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
        public ParagraphPiece create()
        {
            return (ParagraphPiece)base.create();
        }

        public ParagraphPieceStyle setBold(bool bold)
        {
            this.bold = bold;
            return this;
        }
        public ParagraphPieceStyle setItalic(bool italic)
        {
            this.italic = italic;
            return this;
        }
        public ParagraphPieceStyle setUnderline(bool underline)
        {
            this.underline = underline;
            return this;
        }

        /// <summary>
        /// If you know the color code, just to straight to the point! Eg.:
        /// <example>
        /// yellow: FFFF00, black: 000000, red: FF0000, blue: 0000FF, green: 008000, etc...
        /// </example>
        /// If you want, you can use the class Color.whatever_color.
        /// </summary>
        /// <param name="textColor">Hexadecimal color code</param>
        /// <returns></returns>
        public ParagraphPieceStyle setTextColor(string textColor)
        {
            this.textColor = textColor;
            return this;
        }

        public ParagraphPieceStyle setTextColor(Color color)
        {
            this.color = color;
            return this;
        }

        public ParagraphPieceStyle setFont(Font font)
        {
            this.font = font;
            return this;
        }

        /***
        * Pass '50' for something quite big. I am not sure if the unit is pixels or what... 
        * find out later. sorry about that. 
        */
        /// <summary>
        /// Pass '50' for something quite big. I am not sure if the unit is pixels or what... 
        /// find out later. sorry about that. 
        /// </summary>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        public ParagraphPieceStyle setFontSize(string fontSize)
        {
            this.fontSize = fontSize;
            return this;
        }



    }
}