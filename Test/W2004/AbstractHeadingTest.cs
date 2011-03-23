using NUnit.Framework;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Style;

namespace Test.W2004
{
    public class AbstractHeadingTest : Assert
    {
        //### TODO: I won't test the method applyStyle because I will pull this off to another class in order to reuse this for all other "stylable" class. 

        //anonymous implementation - this is the way I leaned how to test abstract classes. 


        [Test]
        public void sanityTest()
        {
            Heading1 heading1 = new Heading1("h111");
            Assert.AreEqual(1, TestUtils.RegexCount(heading1.Template, "<w:p wsp:rsidR*"));
            Assert.AreEqual(1, TestUtils.RegexCount(heading1.Template, "<w:t>[{]value[}]</w:t>"));
                //it has to have the pace holder
            Assert.AreEqual(1, TestUtils.RegexCount(heading1.Template, "</w:p>"));
            Assert.AreEqual(1, TestUtils.RegexCount(heading1.Template, "<w:pStyle w:val=\"Heading1\" />"));
                // it has to replace the Type of Heading

            Assert.NotNull(heading1.Style);
            HeadingStyle style = new HeadingStyle();
            heading1.Style = style;
            Assert.NotNull(heading1.Style);
        }
    }
}