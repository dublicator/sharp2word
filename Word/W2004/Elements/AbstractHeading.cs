using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Heading is utilized to organize documents the same way you do for web pages. 
    /// 
    ///   You can use Heading1 to 3.
    /// </summary>
    /// <typeparam name = "T"></typeparam>
    public class AbstractHeading<T> : IElement, IFluentElementStylable<T> where T : class
    {
        /// <summary>
        /// this is actual heading1, heading2 or heading3.
        /// </summary>
        private readonly string _headingType;

        private readonly string _value; //value/text for the Heading

        protected AbstractHeading(string headingType, string value)
        {
            this._headingType = headingType;
            this._value = value;
        }

        private const string template =
            "\n<w:p wsp:rsidR=\"004429ED\" wsp:rsidRDefault=\"00000000\" wsp:rsidP=\"004429ED\">"
            + "\n	<w:pPr>"
            + "\n		<w:pStyle w:val=\"{heading}\" />"
            + "\n		{styleAlign}"
            + "\n	</w:pPr>"
            + "\n	<w:r>"
            + "\n		{styleText}"
            + "\n		<w:t>{value}</w:t>"
            + "\n	</w:r>"
            + "\n</w:p>";

        private HeadingStyle _style = new HeadingStyle();

        public string Content
        {
            get
            {
                if ("".Equals(this._value))
                {
                    return "";
                }

                //For convention, it should be the last thing before returning the xml content.
                string txt = _style.GetNewContentWithStyle(Template);

                return txt.Replace("{value}", this._value);
            }
        }


        public string Template
        {
            get { return template.Replace("{heading}", this._headingType); }
        }

        public HeadingStyle Style
        {
            get { return _style; }
            set { this._style = _style; }
        }

        //Implements the stylable and the heading classes reuse it
        public T WithStyle()
        {
            this._style.Element = this; //, Heading1.class
            return this._style as T;
        }
    }
}