using Word.Api.Interfaces;

namespace Word.W2004.Elements
{
    /// <summary>
    /// It inserts a Page break at the point you add this class to the document.
    /// </summary>
    public class PageBreak : IElement
    {
        public string Content
        {
            get { return "\n<w:br w:type=\"page\" />"; }
        }

        /// <summary>
        /// This is a different fluent way. 
        /// Because there is no create "with", we will have to create a static method to return an instance of the pageBreak. 
        /// Notice that this class doesn't implement IFluentInterface.
        /// </summary>
        /// <returns></returns>
        public static PageBreak create()
        {
            return new PageBreak();
        }
    }
}