using Word.W2004;

namespace Word.Api.Interfaces
{
    /// <summary>
    ///   The main interface for documents MS Word 2004+.
    /// </summary>
    public interface IDocument : IHasElement
    {
        /// <summary>
        /// </summary>
        /// <value>The URI ready to be added to the document</value>
        string Uri { get; }

        /// <summary>
        /// </summary>
        /// <value>The body of the document</value>
        IBody Body { get; }

        /// <summary>
        /// </summary>
        /// <value>The header that may contain other elements</value>
        IHeader Header { get; }

        /// <summary>
        /// </summary>
        /// <value>The Footer that may contain other elements</value>
        IFooter Footer { get; }

        /// <summary>
        ///   Sets page orientation to Landscape. Default is Portrait
        /// </summary>
        void SetPageOrientationLandscape();

        /// <summary>
        /// Save document to local disk
        /// </summary>
        /// <param name="path">Path to file</param>
        void Save(string path);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title">Represents the title of the document. The title can be different than the file name. The title is used when searching for the document and also when creating Web pages from the document.</param>
        /// <returns>fluent @Document reference</returns>
        Document2004 title(string title);

        /// <summary></summary>
        /// <param name="subject">Represents the subject of the document. This property can be used to group similar files together, so you can search for all files that have the same subject.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 subject(string subject);

        /// <summary></summary>
        /// <param name="keywords">Represents keywords to be used when searching for the document.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 keywords(string keywords);

        /// <summary></summary>
        /// <param name="description">Represents comments to be used when searching for the document.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 description(string description);

        /// <summary></summary>
        /// <param name="category">Represents the author who created the document.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 category(string category);

        /// <summary></summary>
        /// <param name="author">Represents the name of the author of the document.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 author(string author);

        /// <summary></summary>
        /// <param name="lastAuthor">Represents the name of the author who last saved the document.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 lastAuthor(string lastAuthor);

        /// <summary></summary>
        /// <param name="manager">Represents the manager of the author of the document. This property can be used to group similar files together, so you can search for all the files that have the same manager.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 manager(string manager);

        /// <summary></summary>
        /// <param name="company"> Represents the company that employs the author. This property can be used to group similar files together, so you can search for all files that have the same company.</param>
        /// <returns>fluent @Document reference</returns>

        Document2004 company(string company);

        /// <summary>
        /// The encoding you want to use in your document. UTF-8 or ISO8859-1, according to the Enum @Encoding
        /// </summary>
        /// <param name="encoding"></param>
        /// <returns></returns>
        Document2004 encoding(Encoding encoding);
    }
}