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
            IElement par = new Paragraph("");
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTest02()
        {
            IElement par = new Paragraph(new ParagraphPiece(""));
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTest03()
        {
            IElement par = Paragraph.withPieces(new ParagraphPiece(null)).create();
            Assert.AreEqual(par.Content, "");
        }

        [Test]
        public void sanityTestFluent()
        {
            IElement par = Paragraph.with("par01").withStyle().setAlign(Align.CENTER).create();

            basicParagraphCheckings(par, "par01", "center");
        }

        [Test]
        public void testWithStyleBgColor()
        {
            IElement p01 = Paragraph.with("par01").withStyle().setBgColor("FFFF00").create();

            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "FFFF00")); //Background Color

            basicParagraphCheckings(p01, "par01", "left");
        }

        [Test]
        public void testWithStyle()
        {
            IElement p01 = Paragraph.with("").withStyle().create();
            Assert.AreEqual(p01.Content, "");
        }


        [Test]
        public void goodParagraphTest()
        {
            IElement p01 = new Paragraph("This is my paragraph");

            basicParagraphCheckings(p01, "This is my paragraph", null);
        }

        [Test]
        public void testPiecesOne()
        {
            Paragraph p01 = Paragraph.with("Piece01"); //or new Paragraph("xxxx");

            basicParagraphCheckings(p01, "Piece01", null);
        }

        [Test]
        public void testParagraphOneWithStyle()
        {
            Paragraph p01 = (Paragraph)Paragraph.with("111").withStyle().setAlign(Align.CENTER).create();

            basicParagraphCheckings(p01, "111", "center");
        }

        [Test]
        public void testPiecesOneWithStyle()
        {
            ParagraphPiece piece01 = new ParagraphPiece("Piece01");
            Paragraph p01 = new Paragraph(piece01);
            ParagraphPieceStyle style = new ParagraphPieceStyle();
            style.setBold(true);
            style.setItalic(true);
            style.setUnderline(true);
            piece01.setStyle(style);

            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:b/>")); //bold
            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:i/>")); //italic
            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:u w:val=\"single\"/>")); //underline

            basicParagraphCheckings(p01, "Piece01", null);
        }

        [Test]
        public void testPiecesMany()
        {
            ParagraphPiece piece01 = new ParagraphPiece("Piece11111");
            ParagraphPiece piece02 = new ParagraphPiece("Piece22222");

            ParagraphPieceStyle style = new ParagraphPieceStyle();
            style.setBold(true);
            style.setItalic(true);
            piece02.setStyle(style);

            Paragraph p01 = Paragraph.withPieces(piece01, piece02);


            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:t>Piece11111</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:t>Piece22222</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:p wsp:rsidR="));
            Assert.AreEqual(4, TestUtils.regexCount(p01.Content, "<*w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "</w:p>"));

            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:b/>"));
            Assert.AreEqual(1, TestUtils.regexCount(p01.Content, "<w:i/>"));

        }

        [Test]
        public void testEmptyPiecesEtc()
        {
            ParagraphPiece piece01 = new ParagraphPiece("");
            ParagraphPiece piece02 = new ParagraphPiece("");
            Paragraph p01 = new Paragraph(piece01, piece02);
            Assert.AreEqual("", p01.Content);
        }

        [Test]
        public void testEmptyPiecesEtc02()
        {
            ParagraphPiece piece01 = new ParagraphPiece(null);
            ParagraphPiece piece02 = new ParagraphPiece("Piece222");
            Paragraph p01 = new Paragraph(piece01, piece02);
            //Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece222</w:t>"));

            basicParagraphCheckings(p01, "Piece222", null);
        }

        [Test]
        public void testFluent()
        {
            Paragraph p01 = (Paragraph)Paragraph.with("111").withStyle().setAlign(Align.CENTER).create();

            basicParagraphCheckings(p01, "111", "center");
        }

        [Test]
        public void testFluentPiece()
        {
            ParagraphPiece pieces = new ParagraphPiece("111");
            Paragraph p01 = Paragraph.withPieces(pieces);

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

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:t>" + parValue + "</w:t>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:p wsp:rsidR="));
            Assert.AreEqual(2, TestUtils.regexCount(par.Content, "<*w:r>"));
            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "</w:p>"));

            Assert.AreEqual(1, TestUtils.regexCount(par.Content, "<w:jc w:val=\"" + align + "\"/>"));
            Assert.AreEqual(2, TestUtils.regexCount(par.Content, "<*w:pPr>"));

        }
    }
}