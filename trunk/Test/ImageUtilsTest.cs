using NUnit.Framework;
using Word.Utils;
using Word.W2004.Elements;

namespace Test
{
    public class ImageUtilsTest : Assert
    {

    [Test]
    public void SanityTest(){
        Image img = Image.FromFullLocalPath(Util.AppRoot + @"\src\test\resources\dtpick.gif");
        //Image img = new Image(Util.AppRoot + "/src/test/resources/base2logo.png");
        // Image("/Users/leonardo_correa/Desktop/icons_corrup/quote.gif");

        //System.out.println(img.Content);
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
        //for dtPicker.gif
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "R0lGODlhEAAQAPMAAKVNSkpNpUpNSqWmpdbT1v///////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA\nAAAAACH5BAEAAAYALAAAAAAQABAAQwRI0MhJqxmlkLwLyF8hYBpnluJArGzbjkEsB0NtD6PLAjyw\njqeOMANEDVGjm1IJm8WWONLxWDyGQjkdoecjVIOnrzEsKJvPaEEEADs="));

    }

    [Test]
    public void TestLocalImage(){
        Image img = Image.FromFullLocalPath(Util.AppRoot + @"\src\test\resources\dtpick.gif");
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "R0lGODlhEAAQAPMAAKVNSkpNpUpNS"));//just the beginning of...
    }

    [Test]
    public void TestLocalImageFluent(){
        Image img = Image.FromFullLocalPath(Util.AppRoot + @"\src\test\resources\dtpick.gif");
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "R0lGODlhEAAQAPMAAKVNSkpNpUpNS"));//just the beginning of...
    }

    [Test]
    public void TestLocalImageWeb(){
        Image img = Image.FromUrl(Util.AppRoot + @"\src\test\resources\dtpick.gif");
    }

    [Test]
    public void TestLocalImageClasspath(){
        //Image img = Image.From_CLASSPATH(Util.AppRoot + "/src/test/resources/dtpick.gif");
    }

    [Test]
    public void TestLocalImageClasspathFluent(){
        Image img = Image.FromUrl(Util.AppRoot + @"\src\test\resources\dtpick.gif").Create();
    }

    /**
     * ignore because it could fail if you are under a proxy. So for a matter of demostration, uncomment and run it
     */
    [Ignore]
    [Test]
    public void TestWebImage(){
        Image img = Image.FromUrl("http://www.google.com.au/intl/en_com/images/srpr/logo1w.png");
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "width:275pt;height:95pt"));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "BiGQFiipCSS8DCm1Cya1FiyNKzexKTjDDSrLDSvUDi3MEyzHFSvUFC3TGi7bGi/aEi7dGzLcFzPN"));
    }

    [Test]
    public void TestClasspathImage(){
        /*Image img = Image.From_CLASSPATH("/dtpick.gif");

        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "width:16pt;height:16pt"));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "R0lGODlhEAAQAPMAAKVNSkpNpUpNS"));*/
    }

    [Test]
    public void TestDefaultSize() {
        Image img = Image.FromFullLocalPath(Util.AppRoot
                + @"\src\test\resources\base2logo.png");
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content,
                "style=\"width:116pt;height:104pt\""));
    }

    [Test]
    public void TestWidth() {
        Image img = Image.FromFullLocalPath(Util.AppRoot
                + @"\src\test\resources\base2logo.png");
        img.SetWidth(120);
        Assert.AreEqual(0, TestUtils.RegexCount(img.Content,
                "style=\"width:116pt;height:104pt\""));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content,
                "style=\"width:120pt;height:104pt\""));
    }

    [Test]
    public void TestHeight() {
        Image img = Image.FromFullLocalPath(Util.AppRoot
                + @"\src\test\resources\base2logo.png");
        img.SetHeight(110);
        Assert.AreEqual(0, TestUtils.RegexCount(img.Content,
                "style=\"width:116pt;height:104pt\""));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content,
                "style=\"width:116pt;height:110pt\""));
    }

    [Test]
    public void TestWidthAndHeight() {
        Image img = Image.FromFullLocalPath(Util.AppRoot
                + @"\src\test\resources\base2logo.png");
        img.SetWidth(121);
        img.SetHeight(111);
        Assert.AreEqual(0, TestUtils.RegexCount(img.Content,
                "style=\"width:116pt;height:104pt\""));
        Assert.AreEqual(1, TestUtils.RegexCount(img.Content,
                "style=\"width:121pt;height:111pt\""));
    }

    [Test]
    public void TestInvalidImage(){
        Image img = Image.FromFullLocalPath(Util.AppRoot
                + @"\src\test\resources\whatever");
    }
    }
}