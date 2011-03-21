using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    /// <summary>
    /// Heading is utilized to organize documents the same way you do for web pages. 
    /// 
    /// You can use Heading1 to 3. 
    /// </summary>
    /// <typeparam name="E"></typeparam>
    public class AbstractHeading<E> : IElement, IFluentElementStylable<E> where E : class
    {
        /// <summary>
        /// This is actual heading1, heading2 or heading3. 
        /// </summary>
        private string headingType;
        private string value; //value/text for the Heading

        protected AbstractHeading(string headingType, string value)
        {
            this.headingType = headingType;
            this.value = value;
        }

        private string template =
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

        public string getContent()
        {
            if ("".Equals(this.value))
            {
                return "";
            }

            //For convention, it should be the last thing before returning the xml content.
            string txt = style.getNewContentWithStyle(getTemplate());

            return txt.Replace("{value}", this.value);
        }


        public string getTemplate()
        {
            return this.template.Replace("{heading}", this.headingType);
        }
        public HeadingStyle getStyle()
        {
            return style;
        }
        public void setStyle(HeadingStyle style)
        {
            this.style = style;
        }

        /// <summary>
        /// Implements the stylable and the heading classes reuse it
        /// </summary>
        /// <returns></returns>
        public E withStyle()
        {
            this.getStyle().setElement(this); //, Heading1.class
            return this.getStyle() as E;
        }
    }
}