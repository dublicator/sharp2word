using System;
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
		Assert.AreEqual(2, TestUtils.regexCount(bd.Content, "<*w:body>"));
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
		bd.AddEle("<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>"); //this is a breakline
		Assert.AreEqual(2, TestUtils.regexCount(bd.Content, "<*w:body>"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:p wsp:rsidR=\"008979E8\" wsp:rsidRDefault=\"008979E8\"/>"));
	}

	[Test]
	public void testHeader(){
		Body2004 bd = new Body2004();
		bd.Header.AddEle(new Paragraph("header01"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.Content, "<*w:hdr"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:t>header01</w:t>"));
	}
	
	[Test]
	public void testFooter(){
		Body2004 bd = new Body2004();
		bd.Footer.AddEle(new Paragraph("footer01"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.Content, "<*w:ftr"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:t>footer01</w:t>"));
	}
	
	[Test]
	public void testHeaderAndFooterSame(){
		Body2004 bd = new Body2004();
		bd.Header.AddEle(new Paragraph("header01"));
		bd.Footer.AddEle(new Paragraph("footer01"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:t>header01</w:t>"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:t>footer01</w:t>"));
	    Console.WriteLine(bd.Content);
	}
	
	[Test]
	public void testHideHeaderAndFooter(){
		Body2004 bd = new Body2004();
		Assert.False(bd.Header.GetHideHeaderAndFooterFirstPage());// default is false
		bd.Header.SetHideHeaderAndFooterFirstPage(true);
		Assert.True(bd.Header.GetHideHeaderAndFooterFirstPage());
		bd.Header.AddEle(new Paragraph("p1"));
	    Console.WriteLine(bd.Content);
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:t>p1</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.Content, "<*w:sectPr"));
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:hdr w:type=\"first\">"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:ftr w:type=\"first\">"));
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:pgSz w:w"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:pgMar"));
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:cols w:space"));		
	}
	
	[Test]
	public void testshowHeaderAndFooter(){ 
		Body2004 bd = new Body2004();
		bd.Header.AddEle(new Paragraph("p1"));
		Assert.False(bd.Header.GetHideHeaderAndFooterFirstPage());// default is false
		
		Assert.AreEqual(1, TestUtils.regexCount(bd.Content, "<w:t>p1</w:t>"));
		Assert.AreEqual(2, TestUtils.regexCount(bd.Content, "<*w:sectPr"));
		
		Assert.AreEqual(0, TestUtils.regexCount(bd.Content, "<w:hdr w:type=\"first\">"));
		Assert.AreEqual(0, TestUtils.regexCount(bd.Content, "<w:ftr w:type=\"first\">"));
		
		Assert.AreEqual(0, TestUtils.regexCount(bd.Content, "<w:pgSz w:w"));
		Assert.AreEqual(0, TestUtils.regexCount(bd.Content, "<w:pgMar"));
		Assert.AreEqual(0, TestUtils.regexCount(bd.Content, "<w:cols w:space"));		
	} 
    }
}