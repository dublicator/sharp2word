using Word.Api.Interfaces;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class Heading3 : AbstractHeading<HeadingStyle>, IFluentElement<Heading3>
    {
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

        public static Heading3 with(string @string)
        {
            return new Heading3(@string);
        }
    }
}