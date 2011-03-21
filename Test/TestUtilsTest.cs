using NUnit.Framework;
using Word.Utils;

namespace Test
{
    public class TestUtilsTest:Assert
    {
        	[Test]
	public void testRegex(){
		Assert.NotNull(new TestUtils());
		Assert.AreEqual(1, TestUtils.regexCount("leonardo", "leo"));
	}
	
	[Test]
	public void testRegexNotFound(){
		Assert.AreEqual(0, TestUtils.regexCount("xxxxx xxxx", "leo"));
	}
	
	[Test]
	public void testRegexNull01()
    {
		Assert.AreEqual(0, TestUtils.regexCount("xxxxxxxxx", null));
	}
	
	[Test]
	public void testRegexNull02(){
		Assert.AreEqual(0, TestUtils.regexCount(null, "xxx"));
	}
	
	[Test]
	public void testRegexNull03(){
		Assert.AreEqual(0, TestUtils.regexCount(null, null));
	}
    }
}