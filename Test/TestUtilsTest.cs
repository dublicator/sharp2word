using NUnit.Framework;
using Word.Utils;

namespace Test
{
    public class TestUtilsTest:Assert
    {
        	[Test]
	public void testRegex(){
		Assert.NotNull(new TestUtils());
		Assert.AreEqual(1, TestUtils.RegexCount("leonardo", "leo"));
	}
	
	[Test]
	public void testRegexNotFound(){
		Assert.AreEqual(0, TestUtils.RegexCount("xxxxx xxxx", "leo"));
	}
	
	[Test]
	public void testRegexNull01()
    {
		Assert.AreEqual(0, TestUtils.RegexCount("xxxxxxxxx", null));
	}
	
	[Test]
	public void testRegexNull02(){
		Assert.AreEqual(0, TestUtils.RegexCount(null, "xxx"));
	}
	
	[Test]
	public void testRegexNull03(){
		Assert.AreEqual(0, TestUtils.RegexCount(null, null));
	}
    }
}