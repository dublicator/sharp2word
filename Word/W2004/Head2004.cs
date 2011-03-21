using System.Text;
using Word.Api.Interfaces;
using Word.Utils;

namespace Word.W2004
{
    public class Head2004 : IHead
    {
        private StringBuilder content = new StringBuilder("");

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
    }
}