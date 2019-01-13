namespace Exceader.Models
{
    public interface IRow
    {
        ISheet Sheet { get; }
        ICell this[int index] { get; }
        int Index { get; }
    }
}
