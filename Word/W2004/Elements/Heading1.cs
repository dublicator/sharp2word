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

        #region IFluentElement<Heading1> Members

        public Heading1 Create()
        {
            return this;
        }

        #endregion

        /// <summary>
        ///   This method is specific for each class. Constructor can be different...Don't know if we can make it generic
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Heading1 With(string value)
        {
            return new Heading1(value);
        }
    }
}