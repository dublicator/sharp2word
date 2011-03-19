namespace Word.Api.Interfaces
{
    public interface ISuperStylin
    {
        /// <summary>
        /// This method will called by IElement.getContent();
        /// </summary>
        /// <param name="txt"></param>
        /// <returns></returns>
        string getNewContentWithStyle(string txt);

        void setElement(IElement element);

        /// <summary>
        /// This method returns the element. There should be a cast for the return.
        /// The other option to avoid type cast is use covariant type return
        /// </summary>
        /// <returns></returns>
        IElement create();
    }
}