using System;
using NUnit.Framework;
using Word.Utils;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class ImageTest : Assert
    {
        [Test]
        public void sanityTest()
        {
            Image img = new Image(Util.AppRoot + "/src/test/resources/dtpick.gif", ImageLocation.FULL_LOCAL_PATH);
            //Image img = new Image(Utils.getAppRoot() + "/src/test/resources/base2logo.png");
            // Image("/Users/leonardo_correa/Desktop/icons_corrup/quote.gif");

            Console.WriteLine(img.Content);
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
            //for dtPicker.gif
            Assert.AreEqual(1,
                            TestUtils.RegexCount(img.Content,
                                                 "R0lGODlhEAAQAPMAAKVNSkpNpUpNSqWmpdbT1v///////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA\nAAAAACH5BAEAAAYALAAAAAAQABAAQwRI0MhJqxmlkLwLyF8hYBpnluJArGzbjkEsB0NtD6PLAjyw\njqeOMANEDVGjm1IJm8WWONLxWDyGQjkdoecjVIOnrzEsKJvPaEEEADs="));
        }

        [Test]
        public void testLocalImage()
        {
            /*
    	Image img = new Image(Util.getAppRoot() + "/src/test/resources/dtpick.gif", ImageLocation.FULL_LOCAL_PATH);
    	Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*w:pict>"));
    	Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*v:shapetype"));
    	Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*v:shape[ >]")); //white space or >
    	Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "wordml"));
    	Assert.AreEqual(1, TestUtils.regexCount(img.getContent(), "R0lGODlhEAAQAPMAAKVNSkpNpUpNS"));//just the beginning of...*/
        }

        [Test]
        public void testLocalImageFluent()
        {
            /*
        Image img = Image.from_FULL_LOCAL_PATHL(Util.getAppRoot() + "/src/test/resources/dtpick.gif");
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "wordml"));
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(), "R0lGODlhEAAQAPMAAKVNSkpNpUpNS"));//just the beginning of...*/
        }

        [Test]
        public void testLocalImageWeb()
        {
            /*
		Image img = new Image(Util.getAppRoot() + "/src/test/resources/dtpick.gif", ImageLocation.WEB_URL);*/
        }

        [Test]
        public void testLocalImageClasspath()
        {
            /*
    	Image img = new Image(Util.getAppRoot() + "/src/test/resources/dtpick.gif", ImageLocation.CLASSPATH);*/
        }

        [Test]
        public void testLocalImageClasspathFluent()
        {
            //Image img = Image.from_WEB_URL(Utils.getAppRoot() + "/src/test/resources/dtpick.gif").create();
        }

        /**
     * ignore because it could fail if you are under a proxy. So for a matter of demostration, uncomment and run it
     */

        [Ignore]
        [Test]
        public void testWebImage()
        {
            Image img = new Image("http://www.google.com.au/intl/en_com/images/srpr/logo1w.png", ImageLocation.WEB_URL);
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*w:pict>"));
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shapetype"));
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "<*v:shape[ >]")); //white space or >
            Assert.AreEqual(2, TestUtils.RegexCount(img.Content, "wordml"));
            Assert.AreEqual(1, TestUtils.RegexCount(img.Content, "width:275pt;height:95pt"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(img.Content,
                                                 "BiGQFiipCSS8DCm1Cya1FiyNKzexKTjDDSrLDSvUDi3MEyzHFSvUFC3TGi7bGi/aEi7dGzLcFzPN"));
        }

        [Test]
        public void testClasspathImage()
        {
            /* Image img = new Image("/dtpick.gif", ImageLocation.CLASSPATH);

        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*w:pict>"));
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*v:shapetype"));
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "<*v:shape[ >]")); //white space or >
        Assert.AreEqual(2, TestUtils.regexCount(img.getContent(), "wordml"));
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(), "width:16pt;height:16pt"));
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(), "R0lGODlhEAAQAPMAAKVNSkpNpUpNS"));*/
        }

        [Test]
        public void testDefaultSize()
        {
            /*  Image img = new Image(Utils.getAppRoot()
                + "/src/test/resources/base2logo.png", ImageLocation.FULL_LOCAL_PATH);
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(),
                "style=\"width:116pt;height:104pt\""));*/
        }

        [Test]
        public void testWidth()
        {
            /*  Image img = new Image(Utils.getAppRoot()
                + "/src/test/resources/base2logo.png", ImageLocation.FULL_LOCAL_PATH);
        img.setWidth("120");
        Assert.AreEqual(0, TestUtils.regexCount(img.getContent(),
                "style=\"width:116pt;height:104pt\""));
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(),
                "style=\"width:120pt;height:104pt\""));*/
        }

        [Test]
        public void testHeight()
        {
            /*  Image img = new Image(Utils.getAppRoot()
                + "/src/test/resources/base2logo.png", ImageLocation.FULL_LOCAL_PATH);
        img.setHeight("110");
        Assert.AreEqual(0, TestUtils.regexCount(img.getContent(),
                "style=\"width:116pt;height:104pt\""));
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(),
                "style=\"width:116pt;height:110pt\""));*/
        }

        [Test]
        public void testWidthAndHeight()
        {
            /* Image img = new Image(Utils.getAppRoot()
                + "/src/test/resources/base2logo.png", ImageLocation.FULL_LOCAL_PATH);
        img.setWidth("121");
        img.setHeight("111");
        Assert.AreEqual(0, TestUtils.regexCount(img.getContent(),
                "style=\"width:116pt;height:104pt\""));
        Assert.AreEqual(1, TestUtils.regexCount(img.getContent(),
                "style=\"width:121pt;height:111pt\""));*/
        }


        [Test]
        public void testInvalidImage()
        {
            // Image img = new Image(Utils.getAppRoot()
            //       + "/src/test/resources/whatever", ImageLocation.FULL_LOCAL_PATH);
        }
    }
}