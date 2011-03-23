using System;
using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004.Elements
{
    public class Shape:IElement
    {
        private readonly StringBuilder txt = new StringBuilder("");

        private string _tempate =
            "<w:pict>"
            +
            "<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\">"
            + "<v:formulas>"
            + " <v:f eqn=\"sum #0 0 10800\"/>"
            + " <v:f eqn=\"prod #0 2 1\"/>"
            + " <v:f eqn=\"sum 21600 0 @1\"/>"
            + " <v:f eqn=\"sum 0 0 @2\"/>"
            + " <v:f eqn=\"sum 21600 0 @3\"/>"
            + " <v:f eqn=\"if @0 @3 0\"/>"
            + " <v:f eqn=\"if @0 21600 @1\"/>"
            + " <v:f eqn=\"if @0 0 @2\"/>"
            + " <v:f eqn=\"if @0 @4 21600\"/>"
            + " <v:f eqn=\"mid @5 @6\"/>"
            + " <v:f eqn=\"mid @8 @5\"/>"
            + " <v:f eqn=\"mid @7 @8\"/>"
            + " <v:f eqn=\"mid @6 @7\"/>"
            + " <v:f eqn=\"sum @6 0 @5\"/>"
            + " </v:formulas>"
            +
            " <v:path textpathok=\"t\" o:connecttype=\"custom\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" o:connectangles=\"270,180,90,0\"/>"
            + " <v:textpath on=\"t\" fitshape=\"t\"/>"
            + "<v:handles>"
            + " <v:h position=\"#0,bottomRight\" xrange=\"6629,14971\"/>"
            + " </v:handles>"
            + " <o:lock v:ext=\"edit\" text=\"t\" shapetype=\"t\"/>"
            + " </v:shapetype>"
            + "<v:shape id=\"_x0000_i1025\" type=\"#_x0000_t136\" style=\"width:300pt;height:51pt\">"
            + " <v:shadow color=\"#868686\"/>"
            +
            " <v:textpath style=\"font-family:'Arial Black';v-text-kern:t\" trim=\"t\" fitpath=\"t\" string=\"This is WordArt\"/>"
            + " </v:shape>"
            + " </w:pict>";





        #region Implementation of IElement

        /// <summary>
        ///   <p>This method returns the content (XML or HTML) of the Element and the content.</p>
        ///   <p>If you are using W2004, the return will be the XML required to generate the element.</p>
        /// 
        ///   <p>Important: Once you call this method, the Document value is cached an no elements can be added later.</p>
        ///   <example>
        ///     <p>This is the XML that generates a <code>BreakLine</code>:</p>
        ///     <code>
        ///       <w:p wsp:rsidR = '008979E8' wsp:rsidRDefault = '008979E8' />
        ///     </code>
        ///   </example>
        /// </summary>
        /// <value>This is the string value of the element ready to be appended/inserted in the Document.</value>
        public string Content
        {
            get { return _tempate; }
        }

        #endregion
    }
}