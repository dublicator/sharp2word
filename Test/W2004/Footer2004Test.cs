using NUnit.Framework;
using Word.Utils;
using Word.W2004;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class Footer2004Test : Assert
    {
        [Test]
        public void sanityTest()
        {
            Footer2004 ft = new Footer2004();
            Assert.AreEqual("", ft.Content);
        }

        [Test]
        public void testAddEle()
        {
            Footer2004 ft = new Footer2004();
            ft.AddEle(new Paragraph("p1"));
            Assert.AreEqual(2, TestUtils.RegexCount(ft.Content, "<*w:ftr"));
            Assert.AreEqual(1, TestUtils.RegexCount(ft.Content, "<w:t>p1</w:t>"));
            Assert.AreEqual(6, TestUtils.RegexCount(ft.Content, "<w:rStyle w:val=\"PageNumber\"/>"));
        }

        [Test]
        public void testAddEleString()
        {
            Footer2004 ft = new Footer2004();
            ft.AddEle("<w:t>p1</w:t>");
            Assert.AreEqual(2, TestUtils.RegexCount(ft.Content, "<*w:ftr"));
            Assert.AreEqual(1, TestUtils.RegexCount(ft.Content, "<w:t>p1</w:t>"));
        }


        [Test]
        public void testFooterWithNoPageNumber()
        {
            Footer2004 ft = new Footer2004();
            ft.ShowPageNumber(false);
            ft.AddEle(new Paragraph("p1"));
            Assert.AreEqual(2, TestUtils.RegexCount(ft.Content, "<*w:ftr"));
            Assert.AreEqual(1, TestUtils.RegexCount(ft.Content, "<w:t>p1</w:t>"));
            Assert.AreEqual(0, TestUtils.RegexCount(ft.Content, "<w:rStyle w:val=\"PageNumber\"/>"));
        }
    }
}