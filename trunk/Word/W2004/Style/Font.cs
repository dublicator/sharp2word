namespace Word.W2004.Style
{
    public class Font
    {
        private string value;

        Font(string value)
        {
            this.value = value;
        }

        public string getValue()
        {
            return value;
        }

        public static Font COURIER
        {
            get
            {
                return new Font("Courier");
            }
        }

        public static Font CAMBRIA
        {
            get
            {
                return new Font("Cambria");
            }
        }

        public static Font TIMES_NEW_ROMAN
        {
            get
            {
                return new Font("Times New Roman");
            }
        }

        public static Font CALIBRI
        {
            get
            {
                return new Font("Calibri");
            }
        }
    }
}