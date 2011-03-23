using NUnit.Framework;
using Word.Utils;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class PageBreakTest : Assert
    {
        [Test]
        public void testPageBreak()
        {
            PageBreak pb = new PageBreak();
            Assert.AreEqual(1, TestUtils.RegexCount(pb.Content, "<w:br w:type=\"page\" />"));
        }

        [Test]
        public void testPageBreakFluent()
        {
            PageBreak pb = PageBreak.Create();
            Assert.AreEqual(1, TestUtils.RegexCount(pb.Content, "<w:br w:type=\"page\" />"));
        }
    }
}