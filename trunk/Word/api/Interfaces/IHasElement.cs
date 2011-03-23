namespace Word.Api.Interfaces
{
    /// <summary>
    ///   <p>This interface is implemented by <code>elements</code> that can accept others elements to be processed and inserted inside.</p> 
    ///   <p>Eg.: Paragraph can have an Image inside.</p> 
    /// 
    ///   <blockquote>
    ///     <code>
    ///       IDocument myDoc = new Document2004(); 
    ///       myDoc.getBody().addEle(new Paragraph("This document is an example of paragraph: " + new Image("/server_root/dtpick.gif", ImageLocation.FULL_LOCAL_PATH).getContent())); 
    ///     </code>
    ///   </blockquote>
    /// 
    ///   <p>The Element Document is a good example. It HAS TO accepts others elements:</p> 
    /// 
    ///   <blockquote>
    ///     <code>
    ///       IDocument myDoc = new Document2004();
    ///       myDoc.getBody().addEle(new Heading1("My Header1")); 
    ///     </code>
    ///   </blockquote>
    /// </summary>
    public interface IHasElement : IElement
    {
        /// <summary>
        ///   The content of the Element will be evaluated inside the object.
        /// 
        ///   This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name = "e"></param>
        void AddEle(IElement e);

        /// <summary>
        ///   In order to give flexibility, string methods was added.
        ///   Imagine you need to add an element which hasn't been implemented, for example, <b>Graphs</b>.
        ///   You know how to generate the XML to generate a Graph because you figured it out by yourself.
        ///   In this case you could do this:
        ///   <code>
        ///     IDocument myDoc = new Document2004();
        ///     string myXml = myMethodToReturnXML(); //This method return the XML to generate my graph - you have to write this one
        ///     myDoc.getBody().addEle(new Paragraph("This is a my graph: " + myXml));
        ///   </code>
        ///   This is an alias to 'getBody().addEle'
        /// </summary>
        /// <param name = "str"></param>
        void AddEle(string str);
    }
}