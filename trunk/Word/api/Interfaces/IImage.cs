using Word.W2004.Elements;

namespace Word.Api.Interfaces
{
    public interface IImage : IElement
    {
        Image SetWidth(double value);

        Image SetHeight(double value);

        Image Scale(double percent);
    }
}