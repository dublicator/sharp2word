namespace Word.W2004.Style
{
    public class Color
    {
        private string value;

        private Color(string value)
        {
            this.value = value;
        }

        public string getValue()
        {
            return value;
        }

        public static Color YELLOW
        {
            get { return new Color("FFFF00"); }
        }

        public static Color BLACK
        {
            get { return new Color("000000"); }
        }

        public static Color RED
        {
            get { return new Color("FF0000"); }
        }

        public static Color BLUE
        {
            get { return new Color("0000FF"); }
        }

        public static Color GREEN
        {
            get { return new Color("008000"); }
        }
    }
}