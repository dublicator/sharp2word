using NUnit.Framework;
using Word.Utils;

namespace Test
{
    public class TestUtilsTest : Assert
    {
        [Test]
        public void TestRegex()
        {
            Assert.NotNull(new TestUtils());
            Assert.AreEqual(1, TestUtils.RegexCount("leonardo", "leo"));
        }

        [Test]
        public void TestRegexNotFound()
        {
            Assert.AreEqual(0, TestUtils.RegexCount("xxxxx xxxx", "leo"));
        }

        [Test]
        public void TestRegexNull01()
        {
            Assert.AreEqual(0, TestUtils.RegexCount("xxxxxxxxx", null));
        }

        [Test]
        public void TestRegexNull02()
        {
            Assert.AreEqual(0, TestUtils.RegexCount(null, "xxx"));
        }

        [Test]
        public void TestRegexNull03()
        {
            Assert.AreEqual(0, TestUtils.RegexCount(null, null));
        }
    }
}