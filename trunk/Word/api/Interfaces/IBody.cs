namespace Word.Api.Interfaces
{
    public interface IBody : IHasElement
    {
        /*
         *
         * This is to provide another way to add raw text to the body or replace something like style or whatever you want.
         * 
         * Like Serializable, IBody has no methods or fields and serves only to identify the semantics of being IBody. 
         */

        IHeader Header { get; }

        IFooter Footer { get; }
    }
}