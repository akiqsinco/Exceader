namespace Exceader.Models
{
    public interface ICell
    {
        IRow Row { get; }
        string Value { get; }
        bool IsEmpty { get; }
        int Index { get; }
        string Id { get; }
    }
}
