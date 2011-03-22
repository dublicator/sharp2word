using NUnit.Framework;
using Word.Utils;

namespace Test
{
    public class Utils_Test : Assert
    {
        	[Test]
	public void sanityTest() {
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
	public void getAppRootTest() {
		Util utils = new Util();
		Assert.NotNull(utils);		
		//Assert.True(Util.getAppRoot().contains("/j2w-ejb"));
	}
	
	[Test]
	public void readFileTest() {
		/*string res = Utils.readFile(Utils.getAppRoot() + "/src/test/resources/resources.properties");		
		Assert.AreEqual(1, TestUtils.regexCount(res, "reports.servlet.datasource.lookup"));*/
	}
	
	[Test]
	public void readFileTestException() {
		//string res = Util.readFile(Util.getAppRoot() + "/src/test/resources/not_a_file");		
		//Assert.AreEqual(1, TestUtils.regexCount(res, "FileNotFoundException"));
	}

	[Test]
	public void prettyTest01() {
		string str = Util.pretty("<leo><nada></nada></leo>");
        Assert.True(str.Contains("<leo>\r\n  <nada/>"));
		Assert.AreEqual(2, TestUtils.regexCount(str, "<*leo>"));
		Assert.AreEqual(1, TestUtils.regexCount(str, "<nada/>"));
	}
	
	[Test]
	public void prettyTestException() {
		string crap = "<leo><nada></leo>";
		string str = Util.pretty(crap);
		Assert.AreEqual(crap, crap); // the same crap...
	}
    }
}