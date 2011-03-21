using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Breaks lines like when you press enter in your MS Word. 
    ///   You can insert many Breaklines at once. Eg.:
    ///   <code>new BreakLine(3) </code>
    ///   This will insert 3 Breaklines.
    ///   FOR A MATTER OR EXERCICE, THIS CLASS FOLLOWS ALL THOSE CRASY CHECKSTYLE GUIDELINES.
    /// </summary>
    public class BreakLine : IElement, IFluentElement<BreakLine>
    {
        /// <summary>
        ///   Number of repetitions.
        /// </summary>
        private readonly int _times = 1;

        /// <summary>
        ///   Number of break lines you want to add.
        /// </summary>
        /// <param name = "times">Number of breaklines</param>
        public BreakLine(int times)
        {
            this._times = times;
        }

        /// <summary>
        ///   By default, 1 Number of break line when no number is provided.
        /// </summary>
        public BreakLine()
        {
        }

        #region IElement Members

        public string Content
        {
            get
            {
                StringBuilder res = new StringBuilder("");
                applyBreakLineTimes(res);
                return res.ToString();
            }
        }

        #endregion

        #region IFluentElement<BreakLine> Members

        /// <summary>
        ///   Returns the Breakline object.
        /// </summary>
        /// <returns>The object breakline</returns>
        public BreakLine create()
        {
            return this;
        }

        #endregion

        /// <summary>
        ///   Apply the repetition of break lines.
        /// </summary>
        /// <param name = "res">String to be added content</param>
        private void applyBreakLineTimes(StringBuilder res)
        {
            for (int i = 0; i < this._times; i++)
            {
                res.Append("\n<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>");
            }
        }

        /// <summary>
        ///   Created Breaklines according to the number of times provided.
        /// </summary>
        /// <param name = "value">Number of times</param>
        /// <returns>The Breakline object ready to go!</returns>
        public static BreakLine times(int value)
        {
            return new BreakLine(value);
        }
    }
}