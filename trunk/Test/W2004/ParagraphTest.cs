using NUnit.Framework;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Style;

namespace Test.W2004
{
    public class ParagraphTest:Assert
    {
        	[Test]
	public void sanityTest() {
		IElement par = new Paragraph("");
		Assert.AreEqual(par.getContent(), "");
	}
	
	[Test]
	public void sanityTest02() {
		IElement par = new Paragraph(new ParagraphPiece(""));
		Assert.AreEqual(par.getContent(), "");
	}
	
	[Test]
	public void sanityTest03() {
		IElement par = Paragraph.withPieces(new ParagraphPiece(null)).create();
		Assert.AreEqual(par.getContent(), "");
	}

	[Test]
	public void sanityTestFluent() {
		IElement par = Paragraph.with("par01").withStyle().setAlign(Align.CENTER).create();
		
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:t>par01</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(2, TestUtils.regexCount(par.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:p>"));
		
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:jc w:val=\"center\"/>"));
		Assert.AreEqual(2, TestUtils.regexCount(par.getContent(), "<*w:pPr>"));
		
	}
	
	[Test]
	public void testWithStile() {
		IElement par = Paragraph.with("").withStyle().create();
		Assert.AreEqual(par.getContent(), "");
	}
	
	
	[Test]
	public void goodParagraphTest() {
		IElement par = new Paragraph("This is my paragraph");
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:t>This is my paragraph</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(2, TestUtils.regexCount(par.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(par.getContent(), "</w:p>"));
	}
	
	[Test]
	public void testPiecesOne() {
//		ParagraphPiece piece01 = new ParagraphPiece("Piece01");
//		Paragraph p01 = new Paragraph(piece01);
		Paragraph p01 = Paragraph.with("Piece01"); //or new Paragraph("xxxx");
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece01</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "</w:p>"));
	}
	
	[Test]
	public void testParagraphOneWithStyle() {
//		Paragraph p01 = new Paragraph("111");
//		p01.getStyle().setAlign(ParagraphStyle.Align.CENTER);
		Paragraph p01 = (Paragraph) Paragraph.with("111").withStyle().setAlign(Align.CENTER).create();

		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>111</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "</w:p>"));
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:jc w:val=\"center\"/>"));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:pPr>"));
	}
	
	[Test]
	public void testPiecesOneWithStyle() {
		ParagraphPiece piece01 = new ParagraphPiece("Piece01");
		Paragraph p01 = new Paragraph(piece01);
		ParagraphPieceStyle style = new ParagraphPieceStyle();
		style.setBold(true);
		style.setItalic(true);
		style.setUnderline(true);
		piece01.setStyle(style);
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece01</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "</w:p>"));
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:b/>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:i/>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:u w:val=\"single\"/>"));
	}
	
	[Test]
	public void testPiecesMany() {
		ParagraphPiece piece01 = new ParagraphPiece("Piece11111");
		ParagraphPiece piece02 = new ParagraphPiece("Piece22222");
//		Paragraph p01 = new Paragraph(piece01, piece02);
		
		ParagraphPieceStyle style = new ParagraphPieceStyle();
		style.setBold(true);
		style.setItalic(true);
		piece02.setStyle(style);
		
		Paragraph p01 = (Paragraph) Paragraph.withPieces(piece01, piece02);
		
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece11111</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece22222</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(4, TestUtils.regexCount(p01.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "</w:p>"));
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:b/>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:i/>"));
		
	}
	
	[Test]
	public void testEmptyPiecesEtc() {
		ParagraphPiece piece01 = new ParagraphPiece("");
		ParagraphPiece piece02 = new ParagraphPiece("");
		Paragraph p01 = new Paragraph(piece01, piece02);
		Assert.AreEqual("", p01.getContent());		
	}
	
	[Test]
	public void testEmptyPiecesEtc02() {
		ParagraphPiece piece01 = new ParagraphPiece(null);
		ParagraphPiece piece02 = new ParagraphPiece("Piece222");
		Paragraph p01 = new Paragraph(piece01, piece02);
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>Piece222</w:t>"));
	}

	[Test]
	public void testFluent(){
		Paragraph p01 = (Paragraph) Paragraph.with("111").withStyle().setAlign(Align.CENTER).create();
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>111</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "</w:p>"));
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:jc w:val=\"center\"/>"));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:pPr>"));
	}
	
	[Test]
	public void testFluentPiece(){		
		ParagraphPiece pieces = new ParagraphPiece("111");
		Paragraph p01 = (Paragraph) Paragraph.withPieces(pieces);
		
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:p wsp:rsidR="));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "<w:t>111</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:r>"));
		Assert.AreEqual(1, TestUtils.regexCount(p01.getContent(), "</w:p>"));
		
		Assert.AreEqual(2, TestUtils.regexCount(p01.getContent(), "<*w:pPr>"));
	} 
    }
}