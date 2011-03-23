using NUnit.Framework;
using Word.Utils;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class BreakLineTest : Assert
    {
        [Test]
        public void testBreakDefaultTest()
        {
            BreakLine br = new BreakLine();
            Assert.AreEqual(
                    "\n<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>",
                    br.Content);
            int tot = TestUtils
                    .RegexCount(br.Content,
                            "(<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>)");
            Assert.AreEqual(1, tot);
        }

        [Test]
        public void testBreakTimes()
        {
            BreakLine br = BreakLine.SetTimes(3).Create();
            int tot = TestUtils
                    .RegexCount(br.Content,
                            "(<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>)");
            Assert.AreEqual(3, tot);
        }

        [Test]
        public void testBreakNumberConstructor()
        {
            BreakLine br = new BreakLine(1);
            Assert.AreEqual(
                    "\n<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>",
                    br.Content);
            int tot = TestUtils
                    .RegexCount(br.Content,
                            "(<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>)");
            Assert.AreEqual(1, tot);
        }

        [Test]
        public void testBreakNumberConstructor02()
        {
            BreakLine br02 = new BreakLine(2);
            int tot = TestUtils
                .RegexCount(br02.Content,
                            "(<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>)");
            Assert.AreEqual(2, tot);

            BreakLine br04 = new BreakLine(4);
            //Assert.AreEqual(TestUtils.regexCount)
        }
    }
}