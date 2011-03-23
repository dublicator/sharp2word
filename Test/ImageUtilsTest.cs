using NUnit.Framework;
using Test.Properties;
using Word.Utils;
using BufferedImage = System.Drawing.Image;

namespace Test
{
    public class ImageUtilsTest : Assert
    {
        [Test]
        public void sanityTestLocal()
        {
            //ImageUtils imageUtils = new ImageUtils();
            //Assert.NotNull(imageUtils);
            BufferedImage bufferedImage = Resources.dtpick;
            string hexa = ImageUtils.GetImageHexaBase64(bufferedImage, "gif");
            //Assert.AreEqual(1, TestUtils.regexCount(hexa, "R0lGODlhEAAQAPMAAKVNSkpNpUpNSqWmpdbT1v"));
            Assert.AreEqual(1, TestUtils.RegexCount(hexa, "R0lGODlhEAAQAKIAAKVNSkpNpUpNSqWmpdbT1v"));
        }

        [Test]
        public void exceptionTest()
        {
            /*URL url = new URL("http://localhost:8080/ExampleStruts/img/bullshit.gif");
            BufferedImage bufferedImage = ImageIO.read(url);*/
        }
    }
}