namespace Word.W2004
{
    public class Encoding
    {
        private readonly string _value;

        public static Encoding UTF_8
        {
            get
            {
                return new Encoding("UTF-8");
            }
        }

        public static Encoding ISO8859_1
        {
            get
            {
                return new Encoding("ISO8859-1");
            }
        }

        private Encoding(string value)
        {
            //_encoding = System.Text.Encoding.GetEncoding(value);
            _value = value;
        }

        public string Value
        {
            get { return _value; }
        }
    }

}