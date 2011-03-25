using Word.Api.Interfaces;
using Word.W2004.Elements.TableElements;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class Heading2 : AbstractHeading<HeadingStyle>, IFluentElement<Heading2>
    {
        // Make it private - dont remove the constructor!
        public Heading2(string value)
            : base("Heading2", value)
        {
        }

        #region IFluentElement<Heading2> Members

        public Heading2 Create()
        {
            return this;
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value">The value of the paragraph</param>
        /// <returns>the Fluent @Heading2</returns>
        public static Heading2 With(string value)
        {
            return new Heading2(value);
        }
    }
}