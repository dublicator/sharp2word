using System.Text.RegularExpressions;

namespace Word.Utils
{
    public class TestUtils
    {
        public static int regexCount(string text, string regex)
        {
            if (text == null || regex == null)
            {
                throw new System.ArgumentNullException("Can't be null.");
            }
            Regex pattern = new Regex(regex, RegexOptions.Compiled);
            MatchCollection matcher = pattern.Matches(text);
            return matcher.Count;
        }

        public static void createLocalDoc(string myDoc)
        {
            throw new System.NotImplementedException();
        }
    }
}