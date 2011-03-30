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

        /// <summary>
        /// set page margin size in millimeters (set -1 to keep default value) 
        /// </summary>
        /// <param name="top">margin top</param>
        /// <param name="bottom">margin bottom</param>
        /// <param name="left">margin left</param>
        /// <param name="right">margin right</param>
        void SetMarginBody(double top, double bottom, double left, double right); 
    }
}