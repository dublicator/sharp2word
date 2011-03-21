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
        Assert.AreEqual("", tbl01.getContent());
    }

    // ### TH - Table Header ###
    [Test]
    public void testTableWithArray() {
        Table tbl01 = new Table();
        string[] cols = { "aaa", "bbb" };
        tbl01.addTableEle(TableEle.TH, cols);
        Assert.True(tbl01.getContent().Contains("<w:tbl>"));
        Assert.True(tbl01.getContent().Contains("<w:tblPr>"));
        Assert.True(tbl01.getContent().Contains("</w:tblPr>"));

        Assert.True(tbl01.getContent().Contains("<w:tr ")); // TH
        Assert.True(tbl01.getContent().Contains("<w:t>aaa</w:t>")); // TH
        Assert.True(tbl01.getContent().Contains("<w:t>bbb</w:t>")); // TH
        Assert.True(tbl01.getContent().Contains("</w:tr>"));// TH

        Assert.True(tbl01.getContent().Contains("</w:tbl>"));

    }

    [Test]
    public void testCreateTableEmptyTH() {
        Table tbl03 = new Table();
        tbl03.addTableEle(TableEle.TH, null);
        Assert.AreEqual("", tbl03.getContent());

        tbl03.addTableEle(TableEle.TH, "");
        Assert.AreEqual(2, TestUtils.regexCount(tbl03.getContent(), "<*w:tbl>"));
        Assert.AreEqual(1, TestUtils.regexCount(tbl03.getContent(), "<w:r wsp:rsidRPr=\"004374EC\"> "));
        Assert.AreEqual(1, TestUtils.regexCount(tbl03.getContent(), "<w:t></w:t>"));
    }

    [Test]
    public void testTableHeaderNoRepeat() {
        Table tbl = new Table();
        tbl.addTableEle(TableEle.TH, "Name");
        Assert.AreEqual(0, TestUtils.regexCount(tbl.getContent(), "[{]tblHeader[}]"));
    }
    
    [Test]
    public void testTableHeaderWITHRepeatHeader() {
        Table tbl = new Table();
        tbl.setRepeatTableHeaderOnEveryPage();
        
        tbl.addTableEle(TableEle.TH, "Name");
        Assert.AreEqual(0, TestUtils.regexCount(tbl.getContent(), "[{]tblHeader[}]"));
        Assert.AreEqual(2, TestUtils.regexCount(tbl.getContent(), "<*w:trPr>"));
        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "<w:tblHeader/>"));
        
    }
    
    [Test]
    public void testTableDefinition() {
        TableDefinition tbldef = new TableDefinition();
        Assert.AreEqual(1, TestUtils.regexCount(tbldef.getTop(), "<*w:tbl>"));
        Assert.AreEqual(2, TestUtils.regexCount(tbldef.getTop(), "<*w:tblPr>"));

        Assert.Null(tbldef.getMiddle());

        Assert.AreEqual(1, TestUtils.regexCount(tbldef.getBottom(), "</w:tbl>"));
    }

    [Test]
    public void testTableCol() {
        TableCol tblcol = new TableCol();
        Assert.AreEqual(1, TestUtils.regexCount(tblcol.getTop(), "<*w:tr"));
        Assert.AreEqual(2, TestUtils.regexCount(tblcol.getMiddle(), "<*w:tc>"));
        Assert.AreEqual(1, TestUtils.regexCount(tblcol.getMiddle(),
                "<w:t>[{]value[}]</w:t>")); // test placeholder
        Assert.AreEqual(1, TestUtils.regexCount(tblcol.getBottom(), "</w:tr>"));
    }

    [Test]
    public void testTableFooter() {
        TableFooter tblFooter = new TableFooter();
        Assert.AreEqual(1, TestUtils.regexCount(tblFooter.getTop(), "<*w:tr"));
        Assert.AreEqual(2, TestUtils.regexCount(tblFooter.getMiddle(), "<*w:tc>"));
        Assert.AreEqual(1, TestUtils.regexCount(tblFooter.getMiddle(), "<w:b/>")); // In
                                                                                // this
                                                                                // framework,
                                                                                // footer
                                                                                // has
                                                                                // to
                                                                                // be
                                                                                // bold...
        Assert.AreEqual(1, TestUtils.regexCount(tblFooter.getMiddle(),
                "<w:t>[{]value[}]</w:t>")); // test placeholder
        Assert.AreEqual(1, TestUtils.regexCount(tblFooter.getBottom(), "</w:tr>"));
    }

    [Test]
    public void testNull() {
        Table tbl = new Table();
        tbl.addTableEle(TableEle.TABLE_DEF, null);
        Assert.AreEqual("", tbl.getContent());
    }

    [Test]
    public void testEmpty() {
        Table tbl = new Table();
        string[] arr = {};
        tbl.addTableEle(TableEle.TABLE_DEF, arr);
        Assert.AreEqual("", tbl.getContent());
    }

    // ### Full Table!!! ###
    [Test]
    public void testCreateFullTable() {
        Table tbl = new Table();
        tbl.addTableEle(TableEle.TH, "Name", "Salary");

        tbl.addTableEle(TableEle.TD, "Leonardo", "100,000.00");
        tbl.addTableEle(TableEle.TD, "Romario", "1,000,000.00");

        tbl.addTableEle(TableEle.TF, "Total", "1,100,000.00");

        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "<w:tbl>"));

        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "<w:tblPr>"));
        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "</w:tblPr>"));

        Assert.AreEqual(4, TestUtils.regexCount(tbl.getContent(), "<w:tr"));
        Assert.AreEqual(4, TestUtils.regexCount(tbl.getContent(), "</w:tr>"));

        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>Name</w:t>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>Salary</w:t>"));

        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>Leonardo</w:t>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>100,000.00</w:t>"));
        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>Romario</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(),
                "<w:t>1,000,000.00</w:t>"));

        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>Total</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(),
                "<w:t>1,100,000.00</w:t>"));

        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "</w:tbl>"));

        Assert.AreEqual(tbl.getContent(), tbl.getContent());

    }

    [Test]
    public void testCreateEmptyCell() {
        Table tbl = new Table();
        tbl.addTableEle(TableEle.TD, "Leonardo", "");
        
        Assert.AreEqual(1,
                TestUtils.regexCount(tbl.getContent(), "<w:t>Leonardo</w:t>"));
        Assert.AreEqual(1, TestUtils.regexCount(tbl.getContent(), "<w:t></w:t> "));        
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