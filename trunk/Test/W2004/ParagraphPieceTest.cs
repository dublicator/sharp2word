using NUnit.Framework;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Style;

namespace Test.W2004
{
    public class ParagraphPieceTest : Assert
    {
            [Test]
    public void sanityTest() {
        IElement par = new ParagraphPiece("");
        Assert.AreEqual(par.getContent(), "");
    }

    [Test]
    public void sanityTest01() {
        IElement par = new ParagraphPiece(null);
        Assert.AreEqual(par.getContent(), "");
    }

    [Test]
    public void testGetContent() {
        IElement par = ParagraphPiece.with("piece01");

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        // if there is no style, shouldn't have this
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<*w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleALL() {
        IElement par = ParagraphPiece.with("piece01").withStyle().setBold(true)
                .setItalic(true).setUnderline(true).setFontSize("50")
                .setFont(Font.COURIER).setTextColor("008000").create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"50\" />"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleBold() {
        IElement par = ParagraphPiece.with("piece01").withStyle().setBold(true)
                .setItalic(false).setUnderline(false).create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold

        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"(.*)\" />"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleItalic() {
        IElement par = ParagraphPiece.with("piece01").withStyle()
                .setBold(false).setItalic(true).create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic

        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"(.*)\" />"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleUnderline() {
        IElement par = ParagraphPiece.with("piece01").withStyle()
                .setUnderline(true).create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"(.*)\" />"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleFont() {
        IElement par = ParagraphPiece.with("piece01").withStyle()
                .setFont(Font.COURIER).create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"(.*)\" />"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleTextColor() {
        IElement par = ParagraphPiece.with("piece01").withStyle()
                .setBold(false)
                .setItalic(false).setUnderline(false)
                .setTextColor("008000").create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"(.*)\" />"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }

    [Test]
    public void testGetContentWithStyleFontSize() {
        IElement par = ParagraphPiece.with("piece01").withStyle()
                .setBold(false)
                .setItalic(false)
                .setUnderline(false)
                .setFontSize("50").create();

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:r>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(par.getContent(), "<w:t>piece01</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:r>"));

        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:rPr>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:b/>")); // bold
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:i/>")); // italic
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "<w:u w:val=\"single\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(), "w:color w:val=\"008000\"/>")); // underline
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
        Assert.AreEqual(0, TestUtils.regexCount(par.getContent(),
                "<wx:font wx:val=\"Courier\"/>"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:sz w:val=\"(.*)\" />"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(),
                "<w:sz-cs w:val=\"(.*)\" />"));
        Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:rPr>"));
    }  
    }
}