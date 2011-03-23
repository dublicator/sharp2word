using NUnit.Framework;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Elements.TableElements;

namespace Test.W2004
{
    public class TableTest:Assert
    {
            [Test]
    // New table has to return ""
    public void testCreateEmptyTable() {
        Table tbl01 = new Table();
        Assert.AreEqual("", tbl01.Content);
    }

    // ### TH - Table Header ###
    [Test]
    public void testTableWithArray() {
        Table tbl01 = new Table();
        string[] cols = { "aaa", "bbb" };
        tbl01.AddTableEle(TableEle.TH, cols);
        Assert.True(tbl01.Content.Contains("<w:tbl>"));
        Assert.True(tbl01.Content.Contains("<w:tblPr>"));
        Assert.True(tbl01.Content.Contains("</w:tblPr>"));

        Assert.True(tbl01.Content.Contains("<w:tr ")); // TH
        Assert.True(tbl01.Content.Contains("<w:t>aaa</w:t>")); // TH
        Assert.True(tbl01.Content.Contains("<w:t>bbb</w:t>")); // TH
        Assert.True(tbl01.Content.Contains("</w:tr>"));// TH

        Assert.True(tbl01.Content.Contains("</w:tbl>"));

    }

    [Test]
    public void testCreateTableEmptyTH() {
        Table tbl03 = new Table();
        tbl03.AddTableEle(TableEle.TH, null);
        Assert.AreEqual("", tbl03.Content);

        tbl03.AddTableEle(TableEle.TH, "");
        Assert.AreEqual(2, TestUtils.RegexCount(tbl03.Content, "<*w:tbl>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl03.Content, "<w:r wsp:rsidRPr=\"004374EC\"> "));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl03.Content, "<w:t></w:t>"));
    }

    [Test]
    public void testTableHeaderNoRepeat() {
        Table tbl = new Table();
        tbl.AddTableEle(TableEle.TH, "Name");
        Assert.AreEqual(0, TestUtils.RegexCount(tbl.Content, "[{]tblHeader[}]"));
    }
    
    [Test]
    public void testTableHeaderWITHRepeatHeader() {
        Table tbl = new Table();
        tbl.SetRepeatTableHeaderOnEveryPage();
        
        tbl.AddTableEle(TableEle.TH, "Name");
        Assert.AreEqual(0, TestUtils.RegexCount(tbl.Content, "[{]tblHeader[}]"));
        Assert.AreEqual(2, TestUtils.RegexCount(tbl.Content, "<*w:trPr>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content, "<w:tblHeader/>"));
        
    }
    
    [Test]
    public void testTableDefinition() {
        TableDefinition tbldef = new TableDefinition();
        Assert.AreEqual(1, TestUtils.RegexCount(tbldef.Top, "<*w:tbl>"));
        Assert.AreEqual(2, TestUtils.RegexCount(tbldef.Top, "<*w:tblPr>"));

        Assert.Null(tbldef.Middle);

        Assert.AreEqual(1, TestUtils.RegexCount(tbldef.Bottom, "</w:tbl>"));
    }

    [Test]
    public void testTableCol() {
        TableCol tblcol = new TableCol();
        Assert.AreEqual(1, TestUtils.RegexCount(tblcol.Top, "<*w:tr"));
        Assert.AreEqual(2, TestUtils.RegexCount(tblcol.Middle, "<*w:tc>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tblcol.Middle,
                "<w:t>[{]value[}]</w:t>")); // test placeholder
        Assert.AreEqual(1, TestUtils.RegexCount(tblcol.Bottom, "</w:tr>"));
    }

    [Test]
    public void testTableFooter() {
        TableFooter tblFooter = new TableFooter();
        Assert.AreEqual(1, TestUtils.RegexCount(tblFooter.Top, "<*w:tr"));
        Assert.AreEqual(2, TestUtils.RegexCount(tblFooter.Middle, "<*w:tc>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tblFooter.Middle, "<w:b/>")); // In
                                                                                // this
                                                                                // framework,
                                                                                // footer
                                                                                // has
                                                                                // to
                                                                                // be
                                                                                // bold...
        Assert.AreEqual(1, TestUtils.RegexCount(tblFooter.Middle,
                "<w:t>[{]value[}]</w:t>")); // test placeholder
        Assert.AreEqual(1, TestUtils.RegexCount(tblFooter.Bottom, "</w:tr>"));
    }

    [Test]
    public void testNull() {
        Table tbl = new Table();
        tbl.AddTableEle(TableEle.TABLE_DEF, null);
        Assert.AreEqual("", tbl.Content);
    }

    [Test]
    public void testEmpty() {
        Table tbl = new Table();
        string[] arr = {};
        tbl.AddTableEle(TableEle.TABLE_DEF, arr);
        Assert.AreEqual("", tbl.Content);
    }

    // ### Full Table!!! ###
    [Test]
    public void testCreateFullTable() {
        Table tbl = new Table();
        tbl.AddTableEle(TableEle.TH, "Name", "Salary");

        tbl.AddTableEle(TableEle.TD, "Leonardo", "100,000.00");
        tbl.AddTableEle(TableEle.TD, "Romario", "1,000,000.00");

        tbl.AddTableEle(TableEle.TF, "Total", "1,100,000.00");

        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content, "<w:tbl>"));

        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content, "<w:tblPr>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content, "</w:tblPr>"));

        Assert.AreEqual(4, TestUtils.RegexCount(tbl.Content, "<w:tr"));
        Assert.AreEqual(4, TestUtils.RegexCount(tbl.Content, "</w:tr>"));

        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>Name</w:t>"));
        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>Salary</w:t>"));

        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>Leonardo</w:t>"));
        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>100,000.00</w:t>"));
        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>Romario</w:t>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content,
                "<w:t>1,000,000.00</w:t>"));

        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>Total</w:t>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content,
                "<w:t>1,100,000.00</w:t>"));

        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content, "</w:tbl>"));

        Assert.AreEqual(tbl.Content, tbl.Content);

    }

    [Test]
    public void testCreateEmptyCell() {
        Table tbl = new Table();
        tbl.AddTableEle(TableEle.TD, "Leonardo", "");
        
        Assert.AreEqual(1,
                TestUtils.RegexCount(tbl.Content, "<w:t>Leonardo</w:t>"));
        Assert.AreEqual(1, TestUtils.RegexCount(tbl.Content, "<w:t></w:t> "));        
    }
    
//    [Test]
//    public void testTableHeader() {
//        Table tbl = new Table();
//        tbl.addTableEle(TableEle.TD, "Leonardo", "");
//
//        Assert.AreEqual(1,
//                TestUtils.regexCount(tbl.getContent(), "<w:t>Leonardo</w:t>"));
//        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "<w:t></w:t> "));
//    } 
    }
}