namespace Word
{
    public class Encoding
    {
        private string _value;

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

        Encoding(string value)
        {
            this._value = value;
        }

        public string getValue()
        {
            return _value;
        }

    }

}