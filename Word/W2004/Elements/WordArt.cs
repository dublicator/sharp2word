using System.Drawing;
using System.Text;
using Word.Api.Interfaces;
using Word.Utils;
using Word.W2004.Style;

namespace Word.W2004.Elements
{
    public class WordArt : IElement
    {
        private const string tempate =
            "<w:pict>"
            +
            "<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\""
            + "</v:shapetype>"
            + "</w:pict>";

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
                _txt.Replace("{Text}", this._text);
                _txt.Replace("{Font}", this._font.Value);
                _txt.Replace("{Width}", this._width.ToString());
                _txt.Replace("{Height}", this._height.ToString());
                _txt.Replace("{ShadowColor}", ImageUtils.ColorToHex(this._shadowColor));
                _txt.Replace("{FillColor}", ImageUtils.ColorToHex(this._fillColor));
                _txt.Replace("{FontSize}", this._fontsize.ToString());
                _txt.Replace("{Bold}", this._bold ? "font-weight:bold;" : "");
                _txt.Replace("{Italic}", this._italic ? "font-style:italic;" : "");

                string style = "";
                switch (_style)
                {
                    case 1:
                        //style = _style1;
                        break;
                    case 2:
                        //style = _style2;
                        break;
                    case 3:
                        //style = _style3;
                        break;
                }
                _txt.Replace("{Style}", style);

                return _txt.ToString();
            }
        }

        #endregion

        private readonly string _text;

        private readonly StringBuilder _txt = new StringBuilder(tempate);
        private bool _bold;
        private Color _fillColor = Color.Black;
        private WordFont _font = WordFont.CALIBRI;
        private int _fontsize = 36;
        private double _height = 51;
        private bool _italic;
        private Color _shadowColor = Color.White;
        private int _style = 1;

        private double _width = 145.5;

        public WordArt(string text)
        {
            _text = text;
        }

        public WordFont WordFont
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

        public int Fontsize
        {
            get { return _fontsize; }
            set { _fontsize = value; }
        }

        public bool Bold
        {
            get { return _bold; }
            set { _bold = value; }
        }

        public bool Italic
        {
            get { return _italic; }
            set { _italic = value; }
        }

        public int Style
        {
            get { return _style; }
            set { _style = value; }
        }
    }
}