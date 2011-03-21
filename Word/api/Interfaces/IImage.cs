using Word.W2004.Elements;

namespace Word.Api.Interfaces
{
    public interface IImage : IElement
    {
        Image setWidth(string value);
        Image setHeight(string value);
    }
}