namespace Word.Api.Interfaces
{
    public interface IFooter : IHasElement
    {
        void ShowPageNumber(bool value); //default is true  
    }
}