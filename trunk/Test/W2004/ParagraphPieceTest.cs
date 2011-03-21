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
        public void sanityTest()
        {
            IElement par = new ParagraphPiece("");
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTest01()
        {
            IElement par = new ParagraphPiece(null);
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void testGetContent()
        {
            IElement par = ParagraphPiece.with("piece01");

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));

            // if there is no style, shouldn't have this
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<*w:rPr>"));
        }

        [Test]
        public void testGetContentWithStyleALL()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle().setBold(true)
                .setItalic(true).setUnderline(true).setFontSize("50")
                .setFont(Font.COURIER).setTextColor("008000").create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier Bold Italic\" w:h-ansi=\"Courier Bold Italic\"/>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier Bold Italic\"/>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:sz-cs w:val=\"50\" />"));

        }

        [Test]
        public void testGetContentWithStyleBold()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle().setBold(true)
                    .setItalic(false).setUnderline(false).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz-cs w:val=\"(.*)\" />"));

        }
        [Test]
        public void testGetContentWithStyleItalic()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle()
                    .setBold(false).setItalic(true).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:i/>")); // italic

            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz-cs w:val=\"(.*)\" />"));

        }
        [Test]
        public void testGetContentWithStyleUnderline()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle()
                    .setUnderline(true).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic

            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz-cs w:val=\"(.*)\" />"));

        }
        [Test]
        public void testGetContentWithStyleFont()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle()
                    .setFont(Font.COURIER).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz-cs w:val=\"(.*)\" />"));

        }
        [Test]
        public void testGetContentWithStyleTextColor()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle()
                    .setBold(false)
                    .setItalic(false).setUnderline(false)
                    .setTextColor("008000").create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:sz-cs w:val=\"(.*)\" />"));
        }
        [Test]
        public void testGetContentWithStyleFontSize()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle()
                    .setBold(false)
                    .setItalic(false)
                    .setUnderline(false)
                    .setFontSize("50").create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.regexCount(par.Content,
                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content,
                    "<w:sz-cs w:val=\"(.*)\" />"));
        }
        [Test]
        public void testNOsmartFont()
        {
            /***
             * the font is "ARIAL_NARROW", so there should not be any bold tag in it.
             */

            IElement par = ParagraphPiece.with("piece01").withStyle().setFont(Font.ARIAL_NARROW).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));

            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testSmartFontBold()
        {
            /***
             * the font is "ARIAL_NARROW_BOLD", so there has to be a 'smart' bold tag in it.
             * There should not be any 'italic' this time
             */
            IElement par = ParagraphPiece.with("piece01").withStyle().setFont(Font.ARIAL_NARROW_BOLD).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));


            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testSmartFontItalic()
        {
            /***
             * the font is "ARIAL_NARROW_ITALIC", so there has to be a 'smart' Italic tag in it.
             * There should not be any 'bold' this time
             */
            IElement par = ParagraphPiece.with("piece01").withStyle().setFont(Font.ARIAL_NARROW_ITALIC).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                    TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));

            Assert.AreEqual(0, TestUtils.regexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testSmartFontItalicAndBold()
        {
            /***
             * the font is "ARIAL_NARROW_ITALIC", so there has to be both 'smart' Italic and 'bold' tags in it.
             */
            IElement par = ParagraphPiece.with("piece01").withStyle().setFont(Font.ARIAL_NARROW_BOLD_ITALIC).create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testEquivalentSmartFont()
        {
            Paragraph p1 = Paragraph.withPieces(ParagraphPiece.with("same").withStyle().setFont(Font.COURIER).setBold(true).setItalic(true).create());
            Paragraph p2 = Paragraph.withPieces(ParagraphPiece.with("same").withStyle().setFont(Font.COURIER_BOLD_ITALIC).create());

            Assert.True(p1.Content.Equals(p2.Content));
        }

        [Test]
        public void testGetContentWithStyleBGcolor()
        {
            IElement par = ParagraphPiece.with("piece01").withStyle()
            .setBgColor("FFFF00")
            .create();

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "FFFF00")); //Background Color
        }
    }
}