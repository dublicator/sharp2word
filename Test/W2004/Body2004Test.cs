using NUnit.Framework;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004;
using Word.W2004.Elements;

namespace Test.W2004
{
    public class Body2004Test:Assert
    {
        	[Test]
	public void sanityTest(){
		Body2004 bd = new Body2004();
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:body>"));
	}

	[Test]
	public void testAddEle(){
        /*	Body2004 bd = new Body2004();
		bd.addEle(new IElement() {			
			public string getContent() {
				return "I am an element";
			}
		});
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:body>"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "I am an element"));*/
	}

	[Test]
	public void testAaddEleString(){
		Body2004 bd = new Body2004();
		bd.addEle("<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>"); //this is a breakline
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:body>"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>"));
	}

	[Test]
	public void testHeader(){
		Body2004 bd = new Body2004();
		bd.getHeader().addEle(new Paragraph("header01"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:hdr"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:t>header01</w:t>"));
	}
	
	[Test]
	public void testFooter(){
		Body2004 bd = new Body2004();
		bd.getFooter().addEle(new Paragraph("footer01"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:ftr"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:t>footer01</w:t>"));
	}
	
	[Test]
	public void testHeaderAndFooterSame(){
		Body2004 bd = new Body2004();
		bd.getHeader().addEle(new Paragraph("header01"));
		bd.getFooter().addEle(new Paragraph("footer01"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:t>header01</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:t>footer01</w:t>"));
		//System.out.println(bd.getContent());
	}
	
	[Test]
	public void testHideHeaderAndFooter(){
		Body2004 bd = new Body2004();
		Assert.False(bd.getHeader().getHideHeaderAndFooterFirstPage());// default is false
		bd.getHeader().setHideHeaderAndFooterFirstPage(true);
		Assert.True(bd.getHeader().getHideHeaderAndFooterFirstPage());
		bd.getHeader().addEle(new Paragraph("p1"));
		//System.out.println(bd.getContent());
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:t>p1</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:sectPr"));
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:hdr w:type=\"first\">"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:ftr w:type=\"first\">"));
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:pgSz w:w"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:pgMar"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:cols w:space"));		
	}
	
	[Test]
	public void testshowHeaderAndFooter(){ 
		Body2004 bd = new Body2004();
		bd.getHeader().addEle(new Paragraph("p1"));
		Assert.False(bd.getHeader().getHideHeaderAndFooterFirstPage());// default is false
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.getContent(), "<w:t>p1</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.getContent(), "<*w:sectPr"));
		
		Assert.AreEqual(0, TestUtils.regexCount(bd.getContent(), "<w:hdr w:type=\"first\">"));
		Assert.AreEqual(0, TestUtils.regexCount(bd.getContent(), "<w:ftr w:type=\"first\">"));
		
		Assert.AreEqual(0, TestUtils.regexCount(bd.getContent(), "<w:pgSz w:w"));
		Assert.AreEqual(0, TestUtils.regexCount(bd.getContent(), "<w:pgMar"));
		Assert.AreEqual(0, TestUtils.regexCount(bd.getContent(), "<w:cols w:space"));		
	} 
    }
}