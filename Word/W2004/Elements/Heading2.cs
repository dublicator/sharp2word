using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class Heading2 : AbstractHeading<HeadingStyle>, IFluentElement<Heading2>
    {
        public Heading2(string value)
            : base("Heading2", value)
        {
        }

        #region IFluentElement<Heading2> Members

        public Heading2 create()
        {
            return this;
        }

        #endregion

        public static Heading2 with(string @string)
        {
            return new Heading2(@string);
        }
    }
}