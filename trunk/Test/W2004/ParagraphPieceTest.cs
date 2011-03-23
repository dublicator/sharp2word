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
            IElement par = ParagraphPiece.With("");
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTest01()
        {
            IElement par = ParagraphPiece.With(null);
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void testGetContent()
        {
            IElement par = ParagraphPiece.With("piece01");

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));

            // if there is no style, shouldn't have this
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<*w:rPr>"));
        }

        [Test]
        public void testGetContentWithStyleALL()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Bold.Italic.Underline.FontSize(24)
                .Font(Font.COURIER).TextColor("008000").Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier Bold Italic\" w:h-ansi=\"Courier Bold Italic\"/>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier Bold Italic\"/>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:sz-cs w:val=\"48\" />"));
        }

        [Test]
        public void testGetContentWithStyleBold()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Bold.Italic.Underline.Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz-cs w:val=\"(.*)\" />"));
        }

        [Test]
        public void testGetContentWithStyleItalic()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Italic.Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz-cs w:val=\"(.*)\" />"));
        }

        [Test]
        public void testGetContentWithStyleUnderline()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Underline.Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz-cs w:val=\"(.*)\" />"));
        }

        [Test]
        public void testGetContentWithStyleFont()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle()
                .Font(Font.COURIER).Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz-cs w:val=\"(.*)\" />"));
        }

        [Test]
        public void testGetContentWithStyleTextColor()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Italic.Underline
                .TextColor("008000").Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:sz-cs w:val=\"(.*)\" />"));
        }

        [Test]
        public void testGetContentWithStyleFontSize()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Italic.Underline
                .FontSize(50).Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));


            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:u w:val=\"single\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "w:color w:val=\"008000\"/>")); // underline
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<w:rFonts w:ascii=\"Courier\" w:h-ansi=\"Courier\"/>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content,
                                                    "<wx:font wx:val=\"Courier\"/>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:sz w:val=\"(.*)\" />"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content,
                                                    "<w:sz-cs w:val=\"(.*)\" />"));
        }

        [Test]
        public void testNOsmartFont()
        {
            /***
             * the font is "ARIAL_NARROW", so there should not be any bold tag in it.
             */

            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(Font.ARIAL_NARROW).Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testSmartFontBold()
        {
            /***
             * the font is "ARIAL_NARROW_BOLD", so there has to be a 'smart' bold tag in it.
             * There should not be any 'italic' this time
             */
            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(Font.ARIAL_NARROW_BOLD).Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));


            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testSmartFontItalic()
        {
            /***
             * the font is "ARIAL_NARROW_ITALIC", so there has to be a 'smart' Italic tag in it.
             * There should not be any 'bold' this time
             */
            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(Font.ARIAL_NARROW_ITALIC).Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testSmartFontItalicAndBold()
        {
            /***
             * the font is "ARIAL_NARROW_ITALIC", so there has to be both 'smart' Italic and 'bold' tags in it.
             */
            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(Font.ARIAL_NARROW_BOLD_ITALIC).Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1,
                            TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void testEquivalentSmartFont()
        {
            Paragraph p1 =
                Paragraph.WithPieces(
                    ParagraphPiece.With("same").WithStyle().Font(Font.COURIER).Bold.Italic.Create());
            Paragraph p2 =
                Paragraph.WithPieces(ParagraphPiece.With("same").WithStyle().Font(Font.COURIER_BOLD_ITALIC).Create());

            Assert.True(p1.Content.Equals(p2.Content));
        }

        [Test]
        public void testGetContentWithStyleBGcolor()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle()
                .BgColor("FFFF00")
                .Create();

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:t>piece01</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "FFFF00")); //Background Color
        }
    }
}