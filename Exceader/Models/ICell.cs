namespace Exceader.Models
{
    public interface ICell
    {
        IRow Row { get; }
        string Value { get; }
        int RowIndex { get; }
        int ColumnIndex { get; }
        string Id { get; }
    }
}
