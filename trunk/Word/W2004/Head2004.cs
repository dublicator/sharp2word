using System.Text;
using Word.Api.Interfaces;
using Word.Utils;

namespace Word.W2004
{
    public class Head2004 : IHead
    {
        private readonly StringBuilder content = new StringBuilder("");

        #region IHead Members

        public string Content
        {
            get
            {
                if ("".Equals(this.content.ToString()))
                {
                    content.Append("\n" + Util.HEAD2004 + "\n");
                }

                return content.ToString();
            }
        }

        #endregion
    }
}