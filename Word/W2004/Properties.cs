using System;
using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004
{
    public class Properties:IElement
    {
        public string Title{get; set; }
        public string Subject { get; set; }
        public string Keywords { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
        public string Author { get; set; }
        public string LastAuthor { get; set; }
        public string Manager { get; set; }
        public string Company { get; set; }
        public string HyperlinkBase { get; set; }
        public int Revision { get; set; }
        public string PresentationFormat { get; set; }
        public string Guid { get; set; }
        //public string AppName { get; set; }
        public string Version { get; set; }
        public int TotalTime { get; set; }
        public DateTime LastPrinted { get; set; }
        public DateTime Created { get; set; }
        public DateTime LastSaved { get; set; }
        public int Pages { get; set; }
        public int Words { get; set; }
        public int Characters { get; set; }
        public int CharactersWithSpaces { get; set; }
        public int Bytes { get; set; }
        public int Lines { get; set; }
        public int Paragraphs { get; set; }

        private const string _props =
            "    <o:DocumentProperties> \n"
            + "        <o:Title>{Title}</o:Title> \n"
            + "        <o:Subject>{Subject}</o:Subject> \n"
            + "        <o:Keywords>{Keywords}</o:Keywords> \n"
            + "        <o:Description>{Description}</o:Description> \n"
            + "        <o:Category>{Category}</o:Category> \n"
            + "        <o:Author>{Author}</o:Author> \n"
            + "        <o:LastAuthor>{LastAuthor}</o:LastAuthor> \n"
            + "        <o:Manager>{Manager}</o:Manager> \n"
            + "        <o:Company>{Company}</o:Company> \n"
            + "        <o:HyperlinkBase>{HyperlinkBase}</o:HyperlinkBase> \n"
            + "        <o:Revision>{Revision}</o:Revision> \n"
            + "        <o:PresentationFormat>{PresentationFormat}</o:PresentationFormat> \n"
            + "        <o:Guid>{Guid}</o:Guid> \n"
            //+ "        <o:AppName>{AppName}</o:AppName> \n"
            + "        <o:Version>{Version}</o:Version> \n"
            + "        <o:TotalTime>{TotalTime}</o:TotalTime> \n"
            + "        <o:LastPrinted>{LastPrinted}</o:LastPrinted> \n"
            + "        <o:Created>{Created}</o:Created> \n"
            + "        <o:LastSaved>{LastSaved}</o:LastSaved> \n"
            + "        <o:Pages>{Pages}</o:Pages> \n"
            + "        <o:Words>{Words}</o:Words> \n"
            + "        <o:Characters>{Characters}</o:Characters> \n"
            + "        <o:CharactersWithSpaces>{CharactersWithSpaces}</o:CharactersWithSpaces> \n"
            + "        <o:Bytes>{Bytes}</o:Bytes> \n"
            + "        <o:Lines>{Lines}</o:Lines> \n"
            + "        <o:Paragraphs>{Paragraphs}</o:Paragraphs> \n"
            + "    </o:DocumentProperties> ";

        readonly StringBuilder content = new StringBuilder("");

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
                if ("".Equals(this.content.ToString()))
                {
                    content.Append("\n" + _props + "\n");
                    content.Replace("{Title}", Title);
                    content.Replace("{Subject}", Subject);
                    content.Replace("{Keywords}", Keywords);
                    content.Replace("{Description}", Description);
                    content.Replace("{Category}", Category);
                    content.Replace("{Author}", Author);
                    content.Replace("{LastAuthor}", LastAuthor);
                    content.Replace("{Manager}", Manager);
                    content.Replace("{Company}", Company);
                    content.Replace("{HyperlinkBase}", HyperlinkBase);
                    content.Replace("{Revision}", Revision.ToString());
                    content.Replace("{PresentationFormat}", PresentationFormat);
                    content.Replace("{Guid}", Guid);
                    //content.Replace("{AppName}", AppName);
                    content.Replace("{Version}", Version);
                    content.Replace("{TotalTime}", TotalTime.ToString());
                    content.Replace("{LastPrinted}", FormatDate(LastPrinted));
                    content.Replace("{Created}", FormatDate(Created));
                    content.Replace("{LastSaved}", FormatDate(LastSaved));
                    content.Replace("{Pages}", Pages.ToString());
                    content.Replace("{Words}", Words.ToString());
                    content.Replace("{Characters}", Characters.ToString());
                    content.Replace("{CharactersWithSpaces}", CharactersWithSpaces.ToString());
                    content.Replace("{Bytes}", Bytes.ToString());
                    content.Replace("{Lines}", Lines.ToString());
                    content.Replace("{Paragraphs}", Paragraphs.ToString());
                }
                return content.ToString();
            }
        }

        #endregion

        private static string FormatDate(DateTime time)
        {
            return string.Format("{0}-{1}-{2}T{3}:{4}:{5}Z", time.Year, time.Month, time.Day, time.Hour, time.Minute,
                                 time.Second);
        }
    }
}