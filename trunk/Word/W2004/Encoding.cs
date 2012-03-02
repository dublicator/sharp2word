namespace Word.W2004
{
    public class Encoding
    {
        private readonly string _value;

        public static Encoding UTF8
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

        public static Encoding WINDOWS_1251
        {
            get
            {
                //This is suitable for cirrylic text
                return new Encoding("windows-1251");
            }
        }

        public string Value
        {
            get { return _value; }
        }
    }

}