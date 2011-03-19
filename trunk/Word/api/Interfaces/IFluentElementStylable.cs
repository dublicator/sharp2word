namespace Word.Api.Interfaces
{
    /// <summary>
    /// I invented the word "Stylable". Don't try to look up at the dictionary pls.
    /// This will make all Style classes fluent. You are able to write code like:
    /// <example>
    ///     <code>
    ///         Heading1.with("h3333").withStyle().setBold(true);    
    ///     </code>
    /// </example>
    /// </summary>
    /// <typeparam name="S"></typeparam>
    public interface IFluentElementStylable<S>
    {
        /// <summary>
        /// This method returns style for the element. The element knows who is his style class, but the style doesn't.
        /// This method will do this:
        /// <list type="number">
        ///    <item>
        ///         set up the itself to the style class
        ///         <code>this.getStyle().setElement(this); //, Heading1.class</code>
        ///     </item>
        ///     <item>
        ///         Return the style class
        ///         <code>return this.getStyle();</code>
        ///     </item>
        /// </list>
        /// </summary>
        /// <returns></returns>
        S withStyle();
    }
}