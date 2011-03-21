using NUnit.Framework;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Style;

namespace Test.W2004
{
    public class HeadingTest:Assert
    {
        	[Test]
	public void testH1Style(){
		Heading1 h1 = new Heading1("222222");
		h1.getStyle().setBold(true);
		h1.getStyle().setItalic(true);
		
		Assert.AreEqual(2, TestUtils.regexCount(h1.getContent(), "<*w:rPr>"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:jc w:val=\"left\" />")); //default is left
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:b/>")); 
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:i/>")); 
		
	}

	[Test]
	public void testH1(){
		Heading1 h1 = new Heading1("h1");
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:t>h1</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "</w:p>"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:pStyle w:val=\"Heading1\" />"));
	}
	
	[Test]
	public void testH1fluent(){
		Heading1 h1 = Heading1.with("h1").create();
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:t>h1</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "</w:p>"));
		Assert.AreEqual(1, TestUtils.regexCount(h1.getContent(), "<w:pStyle w:val=\"Heading1\" />"));
	}
	
	[Test]
	public void testH2(){
		Heading2 h2 = new Heading2("h2");
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "<w:t>h2</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "</w:p>"));
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "<w:pStyle w:val=\"Heading2\" />"));
	}
	
	[Test]
	public void testH2Fluent(){
		Heading2 h2 = Heading2.with("h2").create();
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "<w:t>h2</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "</w:p>"));
		Assert.AreEqual(1, TestUtils.regexCount(h2.getContent(), "<w:pStyle w:val=\"Heading2\" />"));
	}
	
	[Test]
	public void testH3(){
		Heading3 h3 = new Heading3("h3");
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "<w:t>h3</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "</w:p>"));
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "<w:pStyle w:val=\"Heading3\" />"));
	}
	
	[Test]
	public void testH3Fluent(){
		Heading3 h3 = Heading3.with("h3").create();
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "<w:t>h3</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "</w:p>"));
		Assert.AreEqual(1, TestUtils.regexCount(h3.getContent(), "<w:pStyle w:val=\"Heading3\" />"));
	}

	[Test]
	public void testEmpty(){
		Heading1 h1 = new Heading1("");
		Assert.AreEqual("", h1.getContent());

		Heading2 h2 = new Heading2("");
		Assert.AreEqual("", h2.getContent());
		
		Heading3 h3 = new Heading3("");
		Assert.AreEqual("", h3.getContent());
	}
	
	[Test]
	public void testFluent(){
		Heading1 h1 = (Heading1) Heading1.with("h111").withStyle().setBold(true).setItalic(true).setAlign(Align.CENTER).create();
		Heading2 h2 = (Heading2) Heading2.with("h222").withStyle().setBold(true).setItalic(true).setAlign(Align.CENTER).create();
		Heading3 h3 = (Heading3) Heading3.with("h222").withStyle().setBold(true).setItalic(true).setAlign(Align.CENTER).create();
		verifyStyles(h1); 
		verifyStyles(h2); 
		verifyStyles(h3); 
	}

	private void verifyStyles(IElement e) {
		Assert.AreEqual(2, TestUtils.regexCount(e.getContent(), "<*w:rPr>"));
		Assert.AreEqual(1, TestUtils.regexCount(e.getContent(), "<w:jc w:val=\"center\" />")); //default is left
		Assert.AreEqual(1, TestUtils.regexCount(e.getContent(), "<w:b/>")); 
		Assert.AreEqual(1, TestUtils.regexCount(e.getContent(), "<w:i/>"));
	} 
    }
}