using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class Heading1 : AbstractHeading<HeadingStyle>, IFluentElement<Heading1>
    {
        public Heading1(string value)
            : base("Heading1", value)
        {
        }

        /// <summary>
        /// This method is specific for each class. Constructor can be different...Don't know if we can make it generic
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        public static Heading1 with(string @string)
        {
            return new Heading1(@string);
        }

        public Heading1 create()
        {
            return this;
        }
    }
}