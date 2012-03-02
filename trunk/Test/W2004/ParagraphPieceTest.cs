using NUnit.Framework;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Style;

namespace Test.W2004
{
    public class ParagraphPieceTest : Assert
    {
        private void doBasicChecking(IElement par, string value)
        {
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:t>" + value + "</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:rPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:rPr>"));
        }

        [Test]
        public void SanityTest()
        {
            IElement par = ParagraphPiece.With("");
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void SanityTest01()
        {
            IElement par = ParagraphPiece.With(null);
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void TestGetContent()
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
        public void TestGetContentWithStyleAll()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Bold().Italic().Underline().FontSize(24)
                .Font(WordFont.COURIER).TextColor("008000").Create();

            doBasicChecking(par, "piece01");



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
        public void TestGetContentWithStyleBold()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Bold().Italic().Underline().Create();

            doBasicChecking(par, "piece01");

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
        public void TestGetContentWithStyleItalic()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Italic().Create();

            doBasicChecking(par, "piece01");

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
        public void TestGetContentWithStyleUnderline()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Underline().Create();

            doBasicChecking(par, "piece01");


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
        public void TestGetContentWithStyleFont()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle()
                .Font(WordFont.COURIER).Create();

            doBasicChecking(par, "piece01");


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
        public void TestGetContentWithStyleTextColor()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Italic().Underline()
                .TextColor("008000").Create();

            doBasicChecking(par, "piece01");


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
        public void TestGetContentWithStyleFontSize()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Italic().Underline()
                .FontSize(50).Create();

            doBasicChecking(par, "piece01");


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
        public void TestNOsmartFont()
        {
            /***
             * the font is "ARIAL_NARROW", so there should not be any bold tag in it.
             */

            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(WordFont.ARIAL_NARROW).Create();

            doBasicChecking(par, "piece01");

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void TestSmartFontBold()
        {
            /***
             * the font is "ARIAL_NARROW_BOLD", so there has to be a 'smart' bold tag in it.
             * There should not be any 'italic' this time
             */
            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(WordFont.ARIAL_NARROW_BOLD).Create();

            doBasicChecking(par, "piece01");


            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void TestSmartFontItalic()
        {
            /***
             * the font is "ARIAL_NARROW_ITALIC", so there has to be a 'smart' Italic tag in it.
             * There should not be any 'bold' this time
             */
            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(WordFont.ARIAL_NARROW_ITALIC).Create();

            doBasicChecking(par, "piece01");

            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void TestSmartFontItalicAndBold()
        {
            /***
             * the font is "ARIAL_NARROW_ITALIC", so there has to be both 'smart' Italic and 'bold' tags in it.
             */
            IElement par = ParagraphPiece.With("piece01").WithStyle().Font(WordFont.ARIAL_NARROW_BOLD_ITALIC).Create();

            doBasicChecking(par, "piece01");

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:b/>")); // bold

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:i/>")); // italic
        }

        [Test]
        public void TestEquivalentSmartFont()
        {
            Paragraph p1 =
                Paragraph.WithPieces(
                    ParagraphPiece.With("same").WithStyle().Font(WordFont.COURIER).Bold().Italic().Create());
            Paragraph p2 =
                Paragraph.WithPieces(ParagraphPiece.With("same").WithStyle().Font(WordFont.COURIER_BOLD_ITALIC).Create());

            Assert.True(p1.Content.Equals(p2.Content));
        }

        [Test]
        public void TestGetContentWithStyleBGcolor()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle()
                .BgColor("FFFF00")
                .Create();

            doBasicChecking(par, "piece01");

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "FFFF00")); //Background Color
        }

        [Test]
        public void TestSubscript()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Subscript().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:vertAlign w:val=\"subscript\"/>"));
        }

        [Test]
        public void TestSuperscript()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Superscript().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:vertAlign w:val=\"superscript\"/>"));
        }

        [Test]
        public void TestCaps()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Caps().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:caps/>"));
        }

        [Test]
        public void TestDoubleStrike()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().DoubleStrike().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:dstrike/>"));
        }

        [Test]
        public void TestStrike()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Strike().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:strike/>"));
        }

        [Test]
        public void TestEmboss()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Emboss().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:emboss/>"));
        }

        [Test]
        public void TestImprint()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Imprint().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:imprint/>"));
        }

        [Test]
        public void TestOutline()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Outline().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:outline/>"));
        }

        [Test]
        public void TestShadow()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Shadow().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:shadow/>"));
        }

        [Test]
        public void TestSmallCaps()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().SmallCaps().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:smallCaps/>"));
        }

        [Test]
        public void TestVanish()
        {
            IElement par = ParagraphPiece.With("piece01").WithStyle().Vanish().Create();
            doBasicChecking(par, "piece01");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:vanish/>"));
        }

        [Test]
        public void TestBidiNoValue()
        {
            ParagraphPiece par = ParagraphPiece.With("111");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:t>111</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:r>"));
            Assert.AreEqual(0, TestUtils.RegexCount(par.Content, "<w:lang w:bidi=\".*\" />"));
        }

        [Test]
        public void TestBidi()
        {
            ParagraphPiece par = ParagraphPiece.With("111").WithStyle().Bidi("HE").Create();
            doBasicChecking(par, "111");
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:lang w:bidi=\".*\" />"));

        }
    }
}