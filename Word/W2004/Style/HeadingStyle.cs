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
        private Align _align = Align.LEFT;

        private bool _bold;
        private bool _italic;

        #region ISuperStylin Members

        /// <summary>
        ///   This method holds the logic to replace all place holders for styling.  
        ///   I also noticed if you don't replace the place holder, it doesn't cause any error! 
        ///   But we should try to replace in order to keep the result xml clean.
        /// </summary>
        /// <param name = "txt"></param>
        /// <returns></returns>
        public override string getNewContentWithStyle(string txt)
        {
            string alignValue = "\n            	<w:jc w:val=\"" + _align.Value + "\" />";
            txt = txt.Replace("{styleAlign}", alignValue);
            StringBuilder sbText = new StringBuilder("");

            applyBoldAndItalic(sbText);

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

        private void applyBoldAndItalic(StringBuilder sbText)
        {
            if (this._bold)
            {
                sbText.Append("\n            	<w:b/><w:b-cs/>");
            }
            if (this._italic)
            {
                sbText.Append("\n            	<w:i/>");
            }
        }


        public HeadingStyle setAlign(Align align)
        {
            this._align = align;
            return this;
        }

        public HeadingStyle setBold(bool bold)
        {
            this._bold = bold;
            return this;
        }

        public HeadingStyle setItalic(bool italic)
        {
            this._italic = italic;
            return this;
        }
    }
}