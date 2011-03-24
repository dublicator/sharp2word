using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class Heading1 : AbstractHeading<HeadingStyle>, IFluentElement<Heading1>
    {
        // Make it private - dont remove the constructor!
        public Heading1(string value)
            : base("Heading1", value)
        {
        }

        public Heading1 Create()
        {
            return this;
        }

        /// <summary>
        ///   This method is specific for each class. Constructor can be different...Don't know if we can make it generic
        /// </summary>
        /// <param name = "value">The value of the paragraph</param>
        /// <returns>the Fluent @Heading1</returns>
        public static Heading1 With(string value)
        {
            return new Heading1(value);
        }
    }
}