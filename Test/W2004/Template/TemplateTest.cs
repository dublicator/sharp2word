using System;
using NUnit.Framework;
using Test.Properties;
using Word.Utils;
using Word.W2004.Elements;
using Word.W2004.Elements.TableElements;

namespace Test
{
    public class TemplateTest:Assert
    {
            //@Ignore //just to not break the build for other devs...
    [Test]
    public void testTemplate()
    {
        string xmlTemplate = Resources.ReleaseNotesTemplate.ToString();
        //string xmlTemplate = Util.readFile("src/test/resources/ReleaseNotesTemplate.doc");
        
        xmlTemplate = replacePh(xmlTemplate, "phCompanyName", "EasyWorld - coding for fun pty");
        xmlTemplate = replacePh(xmlTemplate, "phEnv", "Production");
        xmlTemplate = replacePh(xmlTemplate, "phVersion", "1.0 beta");
        xmlTemplate = replacePh(xmlTemplate, "phProjectLeader", "Leonardo Correa");
        
        Table tbl = new Table();
        tbl.addTableEle(TableEle.TH, "Jira Number", "Description");

        tbl.addTableEle(TableEle.TD, "J2W-1234", "Read Templates nicelly");
        tbl.addTableEle(TableEle.TD, "J2W-9999", "Make Java2word funky!!!");

        xmlTemplate = replacePh(xmlTemplate, "phTableIssues", tbl.Content);
        
        Paragraph p01 = Paragraph.with("1) Stop the server").create();
        Paragraph p02 = Paragraph.with("2) Run the script to deploy the app xxx").create();
        Paragraph p03 = Paragraph.with("3) Start the server").create();
        Paragraph p04 = Paragraph.with("4) Hope for the best").create();
        
        string instructions = p01.Content + p02.Content + p03.Content + p04.Content;
        
        //Workaround: phInstructions is already inside a 'text' fragment. 
        //If you know the template, you can remove the whole element and add all Paragraphs
        //* Table above doesn't need workaround because table can be normally inside a paragraph.
        xmlTemplate = replacePh(xmlTemplate, "<w:t>phInstructions</w:t>", instructions); 
        
        xmlTemplate = replacePh(xmlTemplate, "phDateTime", new DateTime().ToString());
        
        //TestUtils.createLocalDoc(xmlTemplate);        
    }
    
    /***
     * Does the Place Holder replacement but LOGS when can not find place holder.
     * @param base Base string that contains the big XML with all placeholders
     * @param placeHolder the actual place holder
     * @param value value to take place
     * @return the new string with place holder replaced
     */
    private static string replacePh(string @base, string placeHolder, string value) {
        if(!@base.Contains(placeHolder)) {
            //don't want to use log now because I want to keep it simple...
            Console.WriteLine("### WARN: couldn't find the place holder: " + placeHolder);
            return @base;
        }        
        return @base.Replace(placeHolder, value);
    }
    }
}