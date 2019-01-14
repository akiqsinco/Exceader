namespace Exceader.Models
{
    public interface ICell
    {
        IRow Row { get; }
        string Value { get; }
        int Index { get; }
        string Id { get; }
        bool IsEmpty { get; }
    }
}
