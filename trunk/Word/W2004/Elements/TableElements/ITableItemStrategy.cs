namespace Word.W2004.Elements.TableElements
{
    /// <summary>
    /// Strategy Design Pattern
    /// </summary>
    public interface ITableItemStrategy
    {
        string getTop();
        string getMiddle();
        string getBottom();
    }
}