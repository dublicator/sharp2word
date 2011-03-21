namespace Word.Api.Interfaces
{
    public interface IHeader : IHasElement
    {
        void setHideHeaderAndFooterFirstPage(bool value);

        bool getHideHeaderAndFooterFirstPage();

        string getHideHeaderAndFooterFirstPageXml();
    }
}