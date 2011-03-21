namespace Word.W2004.Style
{
    public class Align
    {
        public static Align CENTER
        {
            get
            {
                return new Align("center");
            }
        }
        public static Align LEFT
        {
            get
            {
                return new Align("left");
            }
        }
        public static Align RIGHT
        {
            get
            {
                return new Align("right");
            }
        }
        public static Align JUSTIFIED
        {
            get
            {
                return new Align("both");
            }
        }

        private string value;

        Align(string value)
        {
            this.value = value;
        }

        public string getValue()
        {
            return value;
        }
    }
}