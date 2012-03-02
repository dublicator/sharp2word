using NUnit.Framework;
using Word.Utils;
using Word.W2004;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class Header2004Test : Assert
    {
        [Test]
        public void SanityTest()
        {
            Header2004 hd = new Header2004();
            Assert.AreEqual("", hd.Content);
        }

        [Test]
        public void TestAddEle()
        {
            Header2004 hd = new Header2004();
            hd.AddEle(Paragraph.With("p1"));
            Assert.AreEqual(2, TestUtils.RegexCount(hd.Content, "<*w:hdr"));
            Assert.AreEqual(1, TestUtils.RegexCount(hd.Content, "<w:t>p1</w:t>"));
        }

        [Test]
        public void TestAddEleString()
        {
            Header2004 hd = new Header2004();
            hd.AddEle("<w:t>p1</w:t>");
            Assert.AreEqual(2, TestUtils.RegexCount(hd.Content, "<*w:hdr"));
            Assert.AreEqual(1, TestUtils.RegexCount(hd.Content, "<w:t>p1</w:t>"));
        }

        [Test]
        public void TestHideHeaderandFooter()
        {
            //this has to be tested in the body...
            Header2004 hd = new Header2004();
            hd.SetHideHeaderAndFooterFirstPage(true);
            Assert.True(hd.GetHideHeaderAndFooterFirstPage());
        }
    }
}