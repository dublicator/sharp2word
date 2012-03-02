using NUnit.Framework;
using Word.Utils;

namespace Test
{
    public class UtilsTest : Assert
    {
        [Test]
        public void SanityTest()
        {
            /*
		File dir1 = new File(".");
		File dir2 = new File("..");
		try {
			System.out.println("Current dir : " + dir1.getCanonicalPath());
			System.out.println("Parent  dir : " + dir2.getCanonicalPath());
		} catch (Exception e) {
			e.printStackTrace();
		}
		*/
        }

        [Test]
        public void GetAppRootTest()
        {
            Util utils = new Util();
            Assert.NotNull(utils);
            //Assert.True(Util.getAppRoot().contains("/j2w-ejb"));
        }

        [Test]
        public void ReadFileTest()
        {
            /*string res = Utils.readFile(Utils.getAppRoot() + "/src/test/resources/resources.properties");		
		Assert.AreEqual(1, TestUtils.regexCount(res, "reports.servlet.datasource.lookup"));*/
        }

        [Test]
        public void ReadFileTestException()
        {
            //string res = Util.readFile(Util.getAppRoot() + "/src/test/resources/not_a_file");		
            //Assert.AreEqual(1, TestUtils.regexCount(res, "FileNotFoundException"));
        }

        [Test]
        public void PrettyTest01()
        {
            string str = Util.Pretty("<leo><nada></nada></leo>");
            Assert.True(str.Contains("<leo>\r\n  <nada>"));
            Assert.AreEqual(2, TestUtils.RegexCount(str, "<*leo>"));
            Assert.AreEqual(1, TestUtils.RegexCount(str, "</nada>"));
        }

        [Test,Ignore]
        public void PrettyTestException()
        {
            string crap = "<leo><nada></leo>";
            string str = Util.Pretty(crap);
            Assert.AreEqual(crap, crap); // the same crap...
        }
    }
}