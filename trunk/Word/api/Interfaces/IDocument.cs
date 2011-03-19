namespace Word.Api.Interfaces
{
    /// <summary>
    /// The main interface for documents MS Word 2004+.
    /// </summary>
    public interface IDocument : IHasElement
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns>The URI ready to be added to the document</returns>
        string getUri();

        /// <summary>
        /// 
        /// </summary>
        /// <returns>The body of the document</returns>
        IBody getBody();

        /// <summary>
        /// 
        /// </summary>
        /// <returns>The head that contains uri</returns>
        IHead getHead();

        /// <summary>
        /// 
        /// </summary>
        /// <returns>The header that may contain other elements</returns>
        IHeader getHeader();

        /// <summary>
        /// 
        /// </summary>
        /// <returns>The Footer that may contain other elements</returns>
        IFooter getFooter();

        /// <summary>
        /// Sets page orientation to Landscape. Default is Portrait
        /// </summary>
        void setPageOrientationLandscape();
    }
}