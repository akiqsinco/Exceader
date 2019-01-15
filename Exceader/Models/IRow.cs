namespace Exceader.Models
{
    public interface IRow
    {
        ISheet Sheet { get; }
        ICell this[int index] { get; }
        ICell this[string columnName] { get; }
        int Index { get; }
        bool IsEmpty { get; }
    }
}
