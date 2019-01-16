namespace Exceader.Models
{
    public interface ISheet
    {
        IBook Book { get; }
        IRow this[int index] { get; }
        ICell this[int row, int column] { get; }
        ICell this[string cellId] { get; }
        string Name { get; }
        bool IsEmpty { get; }
    }
}
