using Word.Api.Interfaces;

namespace Word.W2004.Style
{
    public abstract class AbstractStyle : ISuperStylin
    {
        private IElement _element;

        #region ISuperStylin Members

        /// <summary>
        ///   This method will called by IElement.getContent();
        /// </summary>
        /// <param name = "txt"></param>
        /// <returns></returns>
        public abstract string GetNewContentWithStyle(string txt);

        public void SetElement(IElement element)
        {
            this._element = element;
        }

        public IElement Create()
        {
            return this._element;
        }

        #endregion
    }
}