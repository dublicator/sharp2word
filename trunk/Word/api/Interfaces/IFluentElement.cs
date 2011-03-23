namespace Word.Api.Interfaces
{
    public interface IFluentElement<F>
    {
        /// <summary>
        ///   This method is just to keep the API consistent when you do something like:
        ///   <example>
        ///     <code>Heading1 myHEading = Heading1.with("h3333").<b>create();</b></code>
        ///     It is just to have that "create()" at the end but the following code is the same:
        ///     <code>Heading1 myHEading = Heading1.with("h3333");</code>
        ///   </example>
        ///   Anyway, it is here for sake of consistency or semantic
        /// </summary>
        /// <returns></returns>
        F Create();
    }
}