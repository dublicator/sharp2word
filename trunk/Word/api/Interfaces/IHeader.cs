namespace Word.Api.Interfaces
{
    public interface IHeader : IHasElement
    {
        void SetHideHeaderAndFooterFirstPage(bool value);

        bool GetHideHeaderAndFooterFirstPage();

        string GetHideHeaderAndFooterFirstPageXml();
    }
}