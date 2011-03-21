namespace Word.W2004.Elements.TableElements
{
    public class TableEle
    {
        private readonly string _value;

        TableEle(string value)
        {
            this._value = value;
        }

        public string Value
        {
            get { return this._value; }
        }

        public static TableEle TH
        {
            get
            {
                return new TableEle("th");
            }
        }

        public static TableEle TF
        {
            get
            {
                return new TableEle("tf");
            }
        }

        public static TableEle TABLE_DEF
        {
            get
            {
                return new TableEle("tableDef");
            }
        }

        public static TableEle TD
        {
            get
            {
                return new TableEle("td");
            }
        }

    }
}