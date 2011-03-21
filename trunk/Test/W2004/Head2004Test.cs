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
            Assert.AreEqual(2, TestUtils.regexCount(hd.getContent(), "<*o:DocumentProperties>")); // open/close test
            Assert.AreEqual(2, TestUtils.regexCount(hd.getContent(), "<*w:fonts>")); // open/close test
            Assert.AreEqual(2, TestUtils.regexCount(hd.getContent(), "<*w:styles>")); // open/close test
            Assert.AreEqual(2, TestUtils.regexCount(hd.getContent(), "<*w:docPr>")); // open/close test
            Assert.AreEqual(1, TestUtils.regexCount(hd.getContent(), "<w:view w:val=\"print\"/>")); // set up as print to be able to view page breaks etc...
        }
    }
}