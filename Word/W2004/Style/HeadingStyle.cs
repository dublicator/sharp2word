using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004.Style
{
    /// <summary>
    ///   Use this class to apply style to the whole Heading1 element.  
    ///   Default align is left
    /// </summary>
    public class HeadingStyle : AbstractStyle, ISuperStylin
    {
        /// <summary>
        ///   Default align is left
        /// </summary>
        private Align _align = Style.Align.LEFT;

        private bool _bold;
        private bool _italic;

        #region ISuperStylin Members

        /// <summary>
        ///   This method holds the logic to Replace all place holders for styling.
        ///   I also noticed if you don't Replace the place holder, it doesn't cause any error!
        ///   But we should try to Replace in order to keep the result xml clean.
        /// </summary>
        /// <param name = "txt"></param>
        /// <returns></returns>
        public override string GetNewContentWithStyle(string txt)
        {
            string alignValue = "\n            	<w:jc w:val=\"" + _align.Value + "\" />";
            txt = txt.Replace("{styleAlign}", alignValue);

            StringBuilder sbText = new StringBuilder("");

            ApplyBoldAndItalic(sbText);

            if (!"".Equals(sbText.ToString()))
            {
                sbText.Insert(0, "\n	 <w:rPr>");
                sbText.Append("\n	 </w:rPr>");
            }

            txt = txt.Replace("{styleText}", sbText.ToString()); //Convention: apply styles
            txt = txt.Replace("[{]style(.*)[}]", ""); //Convention: remove unused styles after...

            return txt;
        }

        #endregion

        private void ApplyBoldAndItalic(StringBuilder sbText)
        {
            if (_bold)
            {
                sbText.Append("\n            	<w:b/><w:b-cs/>");
            }
            if (_italic)
            {
                sbText.Append("\n            	<w:i/>");
            }
        }


        public HeadingStyle SetAlign(Align align)
        {
            _align = align;
            return this;
        }

        /// <summary>
        ///   Heading alignment
        /// </summary>
        /// <param name = "align"></param>
        /// <returns>fluent @HeadingStyle</returns>
        public HeadingStyle Align(Align align)
        {
            _align = align;
            return this;
        }


        public HeadingStyle SetBold(bool bold)
        {
            _bold = bold;
            return this;
        }

        /// <summary>
        ///   Set Heading font to bold
        /// </summary>
        /// <returns>fluent @HeadingStyle</returns>
        public HeadingStyle Bold()
        {
            _bold = true;
            return this;
        }

        public HeadingStyle SetItalic(bool italic)
        {
            _italic = italic;
            return this;
        }

        /// <summary>
        ///   Set Heading font to italic
        /// </summary>
        /// <returns>fluent @HeadingStyle</returns>
        public HeadingStyle Italic()
        {
            _italic = true;
            return this;
        }
    }
}