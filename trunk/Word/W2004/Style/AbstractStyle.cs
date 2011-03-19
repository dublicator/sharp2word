using Word.Api.Interfaces;

namespace Word.W2004.Style
{
    public abstract class AbstractStyle : ISuperStylin
    {
        private IElement _element;

        /// <summary>
        /// This method will called by IElement.getContent();
        /// </summary>
        /// <param name="txt"></param>
        /// <returns></returns>
        public abstract string getNewContentWithStyle(string txt);

        public void setElement(IElement element)
        {
            this._element = element;
        }

        public IElement create()
        {
            return this._element;
        }
        
    }
}