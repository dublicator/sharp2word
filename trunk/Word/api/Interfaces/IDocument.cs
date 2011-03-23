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
        /// <value>The head that contains uri</value>
        IHead Head { get; }

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
    }
}