using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Heading is utilized to organize documents the same way you do for web pages. 
    /// 
    ///   You can use Heading1 to 3.
    /// </summary>
    /// <typeparam name = "E"></typeparam>
    public class AbstractHeading<E> : IElement, IFluentElementStylable<E> where E : class
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

        private HeadingStyle style = new HeadingStyle();

        public string Content
        {
            get
            {
                if ("".Equals(this._value))
                {
                    return "";
                }

                //For convention, it should be the last thing before returning the xml content.
                string txt = style.GetNewContentWithStyle(Template);

                return txt.Replace("{value}", this._value);
            }
        }


        public string Template
        {
            get { return template.Replace("{heading}", this._headingType); }
        }

        public HeadingStyle Style
        {
            get { return style; }
            set { this.style = style; }
        }

        //Implements the stylable and the heading classes reuse it
        public E WithStyle()
        {
            this.style.Element = this; //, Heading1.class
            return this.style as E;
        }
    }
}