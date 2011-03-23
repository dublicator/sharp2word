using System;
using System.Text;
using Word.Api.Interfaces;

namespace Word.W2004
{
    public class Properties : IElement
    {
        private const string _props =
            "    <o:DocumentProperties> "
            + "        <o:Title>{Title}</o:Title> "
            + "        <o:Subject>{Subject}</o:Subject> "
            + "        <o:Keywords>{Keywords}</o:Keywords> "
            + "        <o:Description>{Description}</o:Description> "
            + "        <o:Category>{Category}</o:Category> "
            + "        <o:Author>{Author}</o:Author> "
            + "        <o:LastAuthor>{LastAuthor}</o:LastAuthor> "
            + "        <o:Manager>{Manager}</o:Manager> "
            + "        <o:Company>{Company}</o:Company> "
            + "        <o:HyperlinkBase>{HyperlinkBase}</o:HyperlinkBase> "
            + "        <o:Revision>{Revision}</o:Revision> "
            + "        <o:PresentationFormat>{PresentationFormat}</o:PresentationFormat> "
            + "        <o:Guid>{Guid}</o:Guid> "
            //+ "        <o:AppName>{AppName}</o:AppName> "
            + "        <o:Version>{Version}</o:Version> "
            + "        <o:TotalTime>{TotalTime}</o:TotalTime> "
            + "        <o:LastPrinted>{LastPrinted}</o:LastPrinted> "
            + "        <o:Created>{Created}</o:Created> "
            + "        <o:LastSaved>{LastSaved}</o:LastSaved> "
            + "        <o:Pages>{Pages}</o:Pages> "
            + "        <o:Words>{Words}</o:Words> "
            + "        <o:Characters>{Characters}</o:Characters> "
            + "        <o:CharactersWithSpaces>{CharactersWithSpaces}</o:CharactersWithSpaces> "
            + "        <o:Bytes>{Bytes}</o:Bytes> "
            + "        <o:Lines>{Lines}</o:Lines> "
            + "        <o:Paragraphs>{Paragraphs}</o:Paragraphs> "
            + "    </o:DocumentProperties> ";

        private readonly StringBuilder _content = new StringBuilder("");

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
                if ("".Equals(this._content.ToString()))
                {
                    _content.Append("" + _props + "");
                    _content.Replace("{Title}", Title);
                    _content.Replace("{Subject}", Subject);
                    _content.Replace("{Keywords}", Keywords);
                    _content.Replace("{Description}", Description);
                    _content.Replace("{Category}", Category);
                    _content.Replace("{Author}", Author);
                    _content.Replace("{LastAuthor}", LastAuthor);
                    _content.Replace("{Manager}", Manager);
                    _content.Replace("{Company}", Company);
                    _content.Replace("{HyperlinkBase}", HyperlinkBase);
                    _content.Replace("{Revision}", Revision.ToString());
                    _content.Replace("{PresentationFormat}", PresentationFormat);
                    _content.Replace("{Guid}", Guid);
                    //content.Replace("{AppName}", AppName);
                    _content.Replace("{Version}", Version);
                    _content.Replace("{TotalTime}", TotalTime.ToString());
                    _content.Replace("{LastPrinted}", FormatDate(LastPrinted));
                    _content.Replace("{Created}", FormatDate(Created));
                    _content.Replace("{LastSaved}", FormatDate(LastSaved));
                    _content.Replace("{Pages}", Pages.ToString());
                    _content.Replace("{Words}", Words.ToString());
                    _content.Replace("{Characters}", Characters.ToString());
                    _content.Replace("{CharactersWithSpaces}", CharactersWithSpaces.ToString());
                    _content.Replace("{Bytes}", Bytes.ToString());
                    _content.Replace("{Lines}", Lines.ToString());
                    _content.Replace("{Paragraphs}", Paragraphs.ToString());
                }
                return _content.ToString();
            }
        }

        #endregion

        /// <summary>
        ///   Represents the title of the document. The title can be different than the file name. 
        ///   The title is used when searching for the document and also when creating Web pages from the document.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///   Represents the subject of the document. This property can be used to group similar files together, 
        ///   so you can search for all files that have the same subject.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        ///   Represents keywords to be used when searching for the document.
        /// </summary>
        public string Keywords { get; set; }

        /// <summary>
        ///   Represents comments to be used when searching for the document.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        ///   Represents the category of the document. 
        ///   This property can be used to group similar files together, so you can search for all files that have the same category.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        ///   Represents the author who created the document.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        ///   Represents the name of the author who last saved the document.
        /// </summary>
        public string LastAuthor { get; set; }

        /// <summary>
        ///   Represents the manager of the author of the document. 
        ///   This property can be used to group similar files together, so you can search for all the files that have the same manager.
        /// </summary>
        public string Manager { get; set; }

        /// <summary>
        ///   Represents the company that employs the author. This property can be used to group similar files together, 
        ///   so you can search for all files that have the same company.
        /// </summary>
        public string Company { get; set; }

        /// <summary>
        ///   Represents the path or URL that is used for all hyperlinks with the same 
        ///   base address that are inserted in the document. 
        ///   This can be an Internet address (for example, http://www.microsoft.com), 
        ///   a path to a folder on your hard disk (for example, c:\personal\documents), 
        ///   or a path to a folder on a network (for example, \\myserver\public\documents).
        /// </summary>
        public string HyperlinkBase { get; set; }

        /// <summary>
        ///   Represents the number of times the document has been saved.
        /// </summary>
        public int Revision { get; set; }

        /// <summary>
        ///   Represents the presentation format of the document.
        /// </summary>
        public string PresentationFormat { get; set; }

        /// <summary>
        ///   Represents the globally unique identifier for the document.
        /// </summary>
        public string Guid { get; set; }

        //public string AppName { get; set; }
        /// <summary>
        ///   Represents the version number of the application that created the document.
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        ///   Represents the number of minutes that the document has been open for editing since it was created.
        /// </summary>
        public int TotalTime { get; set; }

        /// <summary>
        ///   Represents the date and time that the document was last printed.
        /// </summary>
        public DateTime LastPrinted { get; set; }

        /// <summary>
        ///   Represents the date and time that the document was originally created.
        /// </summary>
        public DateTime Created { get; set; }

        /// <summary>
        ///   Represents the date and time that the document was last saved.
        /// </summary>
        public DateTime LastSaved { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of pages in the document.
        /// </summary>
        public int Pages { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of words in the document.
        /// </summary>
        public int Words { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of characters in the document.
        /// </summary>
        public int Characters { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of characters (including spaces) in the document.
        /// </summary>
        public int CharactersWithSpaces { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of bytes in the document.
        /// </summary>
        public int Bytes { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of lines in the document.
        /// </summary>
        public int Lines { get; set; }

        /// <summary>
        ///   Represents an estimate of the number of paragraphs in the document.
        /// </summary>
        public int Paragraphs { get; set; }

        private static string FormatDate(DateTime time)
        {
            return string.Format("{0}-{1}-{2}T{3}:{4}:{5}Z", time.Year, time.Month, time.Day, time.Hour, time.Minute,
                                 time.Second);
        }
    }
}