namespace Word.W2004.Elements.TableElements
{
    /// <summary>
    /// Strategy Design Pattern
    /// </summary>
    public interface ITableItemStrategy
    {
        string Top { get; }
        string Middle { get; }
        string Bottom { get; }
    }
}