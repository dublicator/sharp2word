using NUnit.Framework;
using Word.Utils;
using Word.W2004;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class Footer2004Test : Assert
    {
        [Test]
        public void SanityTest()
        {
            Footer2004 ft = new Footer2004();
            Assert.AreEqual("", ft.Content);
        }

        [Test]
        public void TestAddEle()
        {
            Footer2004 ft = new Footer2004();
            ft.AddEle(Paragraph.With("p1"));
            Assert.AreEqual(2, TestUtils.RegexCount(ft.Content, "<*w:ftr"));
            Assert.AreEqual(1, TestUtils.RegexCount(ft.Content, "<w:t>p1</w:t>"));
            Assert.AreEqual(6, TestUtils.RegexCount(ft.Content, "<w:rStyle w:val=\"PageNumber\"/>"));
        }

        [Test]
        public void TestAddEleString()
        {
            Footer2004 ft = new Footer2004();
            ft.AddEle("<w:t>p1</w:t>");
            Assert.AreEqual(2, TestUtils.RegexCount(ft.Content, "<*w:ftr"));
            Assert.AreEqual(1, TestUtils.RegexCount(ft.Content, "<w:t>p1</w:t>"));
        }


        [Test]
        public void TestFooterWithNoPageNumber()
        {
            Footer2004 ft = new Footer2004();
            ft.AddEle(Paragraph.With("p1"));
            Assert.AreEqual(2, TestUtils.RegexCount(ft.Content, "<*w:ftr"));
            Assert.AreEqual(1, TestUtils.RegexCount(ft.Content, "<w:t>p1</w:t>"));
            Assert.AreEqual(0, TestUtils.RegexCount(ft.Content, "<w:rStyle w:val=\"PageNumber\"/>"));
        }
    }
}