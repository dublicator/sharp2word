using NUnit.Framework;
using Word.Utils;
using Word.W2004;

namespace Test.W2004
{
    public class Head2004Test : Assert
    {
        [Test]
        public void testHead()
        {
            Head2004 hd = new Head2004();
            Assert.AreEqual(2, TestUtils.RegexCount(hd.Content, "<*o:DocumentProperties>")); // open/close test
            Assert.AreEqual(2, TestUtils.RegexCount(hd.Content, "<*w:fonts>")); // open/close test
            Assert.AreEqual(2, TestUtils.RegexCount(hd.Content, "<*w:styles>")); // open/close test
            Assert.AreEqual(2, TestUtils.RegexCount(hd.Content, "<*w:docPr>")); // open/close test
            Assert.AreEqual(1, TestUtils.RegexCount(hd.Content, "<w:view w:val=\"print\"/>"));
                // set up as print to be able to view page breaks etc...
        }
    }
}