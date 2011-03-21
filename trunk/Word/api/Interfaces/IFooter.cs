namespace Word.Api.Interfaces
{
    public interface IFooter : IHasElement
    {
        void showPageNumber(bool value); //default is true  
    }
}