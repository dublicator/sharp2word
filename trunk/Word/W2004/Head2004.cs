using System.Text;
using Word.Api.Interfaces;
using Word.Utils;

namespace Word.W2004
{
    public class Head2004 : IHead
    {
        private readonly StringBuilder _content = new StringBuilder("");

        #region IHead Members

        public Properties Properties { get; set; }

        public string Content
        {
            get
            {
                if ("".Equals(this._content.ToString()))
                {
                    _content.Append("\n" + this.Properties.Content + Util.HEAD2004 + "\n");
                }

                return _content.ToString();
            }
        }

        #endregion
    }
}