using Word.W2004.Elements;

namespace Word.Api.Interfaces
{
    public interface IImage : IElement
    {
        Image SetWidth(int value);

        Image SetHeight(int value);
    }
}