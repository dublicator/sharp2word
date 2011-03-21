using System.Text;
using Word.Api.Interfaces;
using Word.W2004.Elements.TableElements;

namespace Word.W2004.Elements
{
    public class Table : IElement
    {
        StringBuilder txt = new StringBuilder("");
        private bool hasBeenCalledBefore = false; // if getContent has already been called, I cached the result for future invocations
        private bool isRepeatTableHeaderOnEveryPage = false;


        public string Content
        {
            get
            {
                if ("".Equals(txt.ToString()))
                {
                    return "";
                }
                if (hasBeenCalledBefore)
                {
                    return txt.ToString();
                }
                else
                {
                    hasBeenCalledBefore = true;
                }

                ITableItemStrategy tableDef = TableFactoryMethod.getInstance().getTableItem(TableEle.TABLE_DEF);

                txt.Insert(0, tableDef.Top);
                txt.Append("\n" + tableDef.Bottom);

                return txt.ToString();
            }
        }

        public void addTableEle(TableEle tableEle, params string[] cols)
        {
            if (cols != null && cols.Length > 0)
            {
                StringBuilder th = new StringBuilder("");

                ITableItemStrategy item = TableFactoryMethod.getInstance().getTableItem(tableEle);

                for (int i = 0; i < cols.Length; i++)
                {
                    //### commented in order to render the cell regardless of null or empty string
                    //if(!"".Equals(cols[i])){ 
                    th.Append("\n" + item.Middle.Replace("{value}", cols[i]));
                    //}
                }

                if (!"".Equals(th.ToString()))
                {
                    th.Insert(0, item.Top);
                    th.Append(item.Bottom);
                }

                string finalResult = setUpRepeatTableHeaderOnEveryPage(th);

                txt.Append(finalResult);//final result appended
            }
        }

        /// <summary>
        /// Adds the repeat header code if this.isRepeatTableHeaderOnEveryPage is true.
        /// If not, it removes {tblHeader} placeholder.
        /// </summary>
        /// <param name="th"></param>
        /// <returns>The final string</returns>
        private string setUpRepeatTableHeaderOnEveryPage(StringBuilder th)
        {
            string res = th.ToString();
            if (this.isRepeatTableHeaderOnEveryPage)
            {
                res = res.Replace("{tblHeader}", "  \n<w:trPr>\n        <w:tblHeader/>\n    </w:trPr>\n ");
            }

            //clean up placeholder  
            res = res.Replace("{tblHeader}", "");
            return res;
        }

        /// <summary>
        /// Pass 'true' if you want to repeat the table header when the table takes more than one page.
        /// Default is false. 
        /// </summary>
        public void setRepeatTableHeaderOnEveryPage()
        {
            this.isRepeatTableHeaderOnEveryPage = true;
        }
    }
}