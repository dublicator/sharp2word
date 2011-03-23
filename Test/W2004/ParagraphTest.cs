using NUnit.Framework;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Style;

namespace Test.W2004
{
    public class ParagraphTest : Assert
    {
        [Test]
        public void sanityTest()
        {
            IElement par = Paragraph.With("");
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTest02()
        {
            IElement par = new Paragraph(ParagraphPiece.With("").Create());
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTest03()
        {
            IElement par = Paragraph.WithPieces(ParagraphPiece.With(null)).Create();
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTestFluent()
        {
            IElement par = Paragraph.With("par01").WithStyle().Align(Align.CENTER).Create();

            basicParagraphCheckings(par, "par01", "center");
        }

        [Test]
        public void testTab()
        {
            Paragraph p01 = Paragraph.WithPieces(
                ParagraphPiece.With("Bloc 1 Price :").WithStyle().Font(Font.CALIBRI).FontSize(2*11).Create(),
                ParagraphPiece.With(" \t 3 200,00 $").WithStyle().Font(Font.CALIBRI).FontSize(2*11).Create()
                ).AddTab(Paragraph.TabAlign.RIGHT, 8931).Create();

            Assert.AreEqual(2, TestUtils.RegexCount(p01.Content, "<w:pPr>"));
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:tabs>"));
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "</w:tabs>"));
            Assert.AreEqual(2, TestUtils.RegexCount(p01.Content, "</w:pPr>"));
        }

        [Test]
        public void testWithStyleBgColor()
        {
            IElement p01 = Paragraph.With("par01").WithStyle().BgColor("FFFF00").Create();

            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "FFFF00")); //Background Color

            basicParagraphCheckings(p01, "par01", "left");
        }

        [Test]
        public void testWithStyle()
        {
            IElement p01 = Paragraph.With("").WithStyle().Create();
            Assert.AreEqual(p01.Content, "");
        }


        [Test]
        public void goodParagraphTest()
        {
            IElement p01 = Paragraph.With("This is my paragraph");

            basicParagraphCheckings(p01, "This is my paragraph", null);
        }

        [Test]
        public void testPiecesOne()
        {
            Paragraph p01 = Paragraph.With("Piece01"); //or new Paragraph("xxxx");

            basicParagraphCheckings(p01, "Piece01", null);
        }

        [Test]
        public void testParagraphOneWithStyle()
        {
            Paragraph p01 = (Paragraph) Paragraph.With("111").WithStyle().Align(Align.CENTER).Create();

            basicParagraphCheckings(p01, "111", "center");
        }

        [Test]
        public void testPiecesOneWithStyle()
        {
            ParagraphPiece piece01 = ParagraphPiece.With("Piece01").WithStyle().Bold.Italic.Underline.Create();
            Paragraph p01 = new Paragraph(piece01);

            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:b/>")); //bold
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:i/>")); //italic
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:u w:val=\"single\"/>")); //underline

            basicParagraphCheckings(p01, "Piece01", null);
        }

        [Test]
        public void testPiecesMany()
        {
            ParagraphPiece piece01 = ParagraphPiece.With("Piece11111");
            ParagraphPiece piece02 = ParagraphPiece.With("Piece22222").WithStyle().Bold.Italic.Create();

            Paragraph p01 = Paragraph.WithPieces(piece01, piece02);


            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:t>Piece11111</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:t>Piece22222</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:p wsp:rsidR="));
            Assert.AreEqual(4, TestUtils.RegexCount(p01.Content, "<*w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "</w:p>"));

            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:b/>"));
            Assert.AreEqual(1, TestUtils.RegexCount(p01.Content, "<w:i/>"));
        }

        [Test]
        public void testEmptyPiecesEtc()
        {
            ParagraphPiece piece01 = ParagraphPiece.With("");
            ParagraphPiece piece02 = ParagraphPiece.With("");

            Paragraph p01 = new Paragraph(piece01, piece02);
            Assert.AreEqual("", p01.Content);
        }

        [Test]
        public void testEmptyPiecesEtc02()
        {
            ParagraphPiece piece01 = ParagraphPiece.With(null);
            ParagraphPiece piece02 = ParagraphPiece.With("Piece222");
            Paragraph p01 = new Paragraph(piece01, piece02);
            //Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece222</w:t>"));

            basicParagraphCheckings(p01, "Piece222", null);
        }

        [Test]
        public void testFluent()
        {
            Paragraph p01 = (Paragraph) Paragraph.With("111").WithStyle().Align(Align.CENTER).Create();

            basicParagraphCheckings(p01, "111", "center");
        }

        [Test]
        public void testFluentPiece()
        {
            ParagraphPiece pieces = ParagraphPiece.With("111");
            Paragraph p01 = Paragraph.WithPieces(pieces);

            basicParagraphCheckings(p01, "111", null);
        }

        /***
         * Verifies, for ONE paragraphPiece, basec values present
        * TODO: Add verification params for: bold, italic, bgcolor, underline
        * @param par the actual paragraph object
        * @param parValue the expected text deiplayed in the paragraph
        * @param align the expected align
        */

        private void basicParagraphCheckings(IElement par, string parValue, string align)
        {
            if (align == null || "".Equals(align))
            {
                align = "left"; //the default
            }

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:t>" + parValue + "</w:t>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:p wsp:rsidR="));
            Assert.AreEqual(2, TestUtils.RegexCount(par.Content, "<*w:r>"));
            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "</w:p>"));

            Assert.AreEqual(1, TestUtils.RegexCount(par.Content, "<w:jc w:val=\"" + align + "\"/>"));
            Assert.AreEqual(2, TestUtils.RegexCount(par.Content, "<*w:pPr>"));
        }
    }
}