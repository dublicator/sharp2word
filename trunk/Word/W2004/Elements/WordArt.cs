using System.Drawing;
using System.Text;
using Word.Api.Interfaces;
using Word.Utils;
using Font = Word.W2004.Style.Font;

namespace Word.W2004.Elements
{
    public class WordArt : IElement
    {
        private const string tempate =
            "\n<w:pict>"
            +"\t<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\">\n"
            + "\t\t<v:formulas>\n"
            + "\t\t\t<v:f eqn=\"sum #0 0 10800\"/>\n"
            + "\t\t\t<v:f eqn=\"prod #0 2 1\"/>\n"
            + "\t\t\t<v:f eqn=\"sum 21600 0 @1\"/>\n"
            + "\t\t\t<v:f eqn=\"sum 0 0 @2\"/>\n"
            + "\t\t\t<v:f eqn=\"sum 21600 0 @3\"/>\n"
            + "\t\t\t<v:f eqn=\"if @0 @3 0\"/>\n"
            + "\t\t\t<v:f eqn=\"if @0 21600 @1\"/>\n"
            + "\t\t\t<v:f eqn=\"if @0 0 @2\"/>\n"
            + "\t\t\t<v:f eqn=\"if @0 @4 21600\"/>\n"
            + "\t\t\t<v:f eqn=\"mid @5 @6\"/>\n"
            + "\t\t\t<v:f eqn=\"mid @8 @5\"/>\n"
            + "\t\t\t<v:f eqn=\"mid @7 @8\"/>\n"
            + " <v:f eqn=\"mid @6 @7\"/>\n"
            + " <v:f eqn=\"sum @6 0 @5\"/>\n"
            + " </v:formulas>\n"
            +" <v:path textpathok=\"t\" o:connecttype=\"custom\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" o:connectangles=\"270,180,90,0\"/>\n"
            + " <v:textpath on=\"t\" fitshape=\"t\"/>\n"
            + "<v:handles>\n"
            + " <v:h position=\"#0,bottomRight\" xrange=\"6629,14971\"/>\n"
            + " </v:handles>\n"
            + " <o:lock v:ext=\"edit\" text=\"t\" shapetype=\"t\"/>\n"
            + " </v:shapetype>\n"
            +"<v:shape id=\"_x0000_i1025\" type=\"#_x0000_t136\" style=\"width:{Width}pt;height:{Height}pt\" fillcolor=\"#{FillColor}\">\n"
            + " <v:shadow color=\"#{ShadowColor}\"/>\n"
            +" <v:textpath style=\"font-family:'{Font}';v-text-kern:t\" trim=\"t\" fitpath=\"t\" string=\"{Text}\"/>\n"
            + " </v:shape>\n"
            + " </w:pict>\n";

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
            get
            {
                _txt.Replace("{Text}", this.Text);
                _txt.Replace("{Font}", this.Font.Value);
                _txt.Replace("{Width}", this.Width.ToString());
                _txt.Replace("{Height}", this.Height.ToString());
                _txt.Replace("{ShadowColor}", ImageUtils.ColorToHex(this.ShadowColor));
                _txt.Replace("{FillColor}", ImageUtils.ColorToHex(this.FillColor));
                return _txt.ToString();
            }
        }

        #endregion

        private readonly StringBuilder _txt = new StringBuilder(tempate);
        private Color _fillColor = Color.Black;
        private Font _font = Font.CALIBRI;
        private double _height = 51;
        private Color _shadowColor = Color.White;

        private double _width = 145.5;

        public string Text { get; set; }

        public Font Font
        {
            get { return _font; }
            set { _font = value; }
        }

        public double Width
        {
            get { return _width; }
            set { _width = value; }
        }

        public double Height
        {
            get { return _height; }
            set { _height = value; }
        }

        public Color ShadowColor
        {
            get { return _shadowColor; }
            set { _shadowColor = value; }
        }

        public Color FillColor
        {
            get { return _fillColor; }
            set { _fillColor = value; }
        }
    }
}