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
		h1.Style.SetBold(true);
		h1.Style.SetItalic(true);
		
		Assert.AreEqual(2, TestUtils.RegexCount(h1.Content, "<*w:rPr>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:jc w:val=\"left\" />")); //default is left
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:b/>")); 
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:i/>")); 
		
	}

	[Test]
	public void testH1(){
		Heading1 h1 = new Heading1("h1");
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:t>h1</w:t>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "</w:p>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:pStyle w:val=\"Heading1\" />"));
	}
	
	[Test]
	public void testH1fluent(){
		Heading1 h1 = Heading1.With("h1").Create();
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:t>h1</w:t>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "</w:p>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h1.Content, "<w:pStyle w:val=\"Heading1\" />"));
	}
	
	[Test]
	public void testH2(){
		Heading2 h2 = new Heading2("h2");
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "<w:t>h2</w:t>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "</w:p>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "<w:pStyle w:val=\"Heading2\" />"));
	}
	
	[Test]
	public void testH2Fluent(){
		Heading2 h2 = Heading2.With("h2").Create();
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "<w:t>h2</w:t>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "</w:p>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h2.Content, "<w:pStyle w:val=\"Heading2\" />"));
	}
	
	[Test]
	public void testH3(){
		Heading3 h3 = new Heading3("h3");
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "<w:t>h3</w:t>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "</w:p>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "<w:pStyle w:val=\"Heading3\" />"));
	}
	
	[Test]
	public void testH3Fluent(){
		Heading3 h3 = Heading3.With("h3").Create();
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "<w:p wsp:rsidR*"));
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "<w:t>h3</w:t>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "</w:p>"));
		Assert.AreEqual(1, TestUtils.RegexCount(h3.Content, "<w:pStyle w:val=\"Heading3\" />"));
	}

	[Test]
	public void testEmpty(){
		Heading1 h1 = new Heading1("");
		Assert.AreEqual("", h1.Content);

		Heading2 h2 = new Heading2("");
		Assert.AreEqual("", h2.Content);
		
		Heading3 h3 = new Heading3("");
		Assert.AreEqual("", h3.Content);
	}
	
	[Test]
	public void testFluent(){
		Heading1 h1 = (Heading1) Heading1.With("h111").WithStyle().SetBold(true).SetItalic(true).SetAlign(Align.CENTER).Create();
		Heading2 h2 = (Heading2) Heading2.With("h222").WithStyle().SetBold(true).SetItalic(true).SetAlign(Align.CENTER).Create();
		Heading3 h3 = (Heading3) Heading3.With("h222").WithStyle().SetBold(true).SetItalic(true).SetAlign(Align.CENTER).Create();
		verifyStyles(h1); 
		verifyStyles(h2); 
		verifyStyles(h3); 
	}

	private void verifyStyles(IElement e) {
		Assert.AreEqual(2, TestUtils.RegexCount(e.Content, "<*w:rPr>"));
		Assert.AreEqual(1, TestUtils.RegexCount(e.Content, "<w:jc w:val=\"center\" />")); //default is left
		Assert.AreEqual(1, TestUtils.RegexCount(e.Content, "<w:b/>")); 
		Assert.AreEqual(1, TestUtils.RegexCount(e.Content, "<w:i/>"));
	} 
    }
}