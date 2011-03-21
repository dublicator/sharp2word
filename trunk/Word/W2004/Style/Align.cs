namespace Word.W2004.Style
{
    public class Align
    {
        private readonly string value;

        private Align(string value)
        {
            this.value = value;
        }

        public static Align CENTER
        {
            get { return new Align("center"); }
        }

        public static Align LEFT
        {
            get { return new Align("left"); }
        }

        public static Align RIGHT
        {
            get { return new Align("right"); }
        }

        public static Align JUSTIFIED
        {
            get { return new Align("both"); }
        }

        public string Value
        {
            get { return value; }
        }
    }
}