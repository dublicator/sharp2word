using Word.Api.Interfaces;
using Word.W2004.Elements.TableElements;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class Heading3 : AbstractHeading<HeadingStyle>, IFluentElement<Heading3>
    {
        // Make it private - dont remove the constructor!
        public Heading3(string value)
            : base("Heading3", value)
        {
        }

        #region IFluentElement<Heading3> Members

        public Heading3 Create()
        {
            return this;
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value">The value of the paragraph</param>
        /// <returns>the Fluent @Heading3</returns>
        public static Heading3 With(string value)
        {
            return new Heading3(value);
        }
    }
}