using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Word.Utils
{
    public class TestUtils
    {
        public static int RegexCount(string text, string regex)
        {
            if (text == null || regex == null)
            {
                throw new System.ArgumentNullException("Can't be null.");
            }
            Regex pattern = new Regex(regex, RegexOptions.Compiled);
            MatchCollection matcher = pattern.Matches(text);
            return matcher.Count;
        }

        public static void CreateLocalDoc(string myDoc)
        {
            string doc = Util.AppRoot + "\\Sharp2word_allInOne.doc";

            using (FileStream fs = new FileStream(doc, FileMode.Create))
            {
                using (StreamWriter w = new StreamWriter(fs, System.Text.Encoding.UTF8))
                {
                    w.Write(Util.Pretty(myDoc));
                }
            }
        }
    }
}