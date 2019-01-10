using System.Collections.Generic;

namespace Exceader.Models
{
    public interface ISheet : IReadOnlyList<IRow>
    {
        IBook Book { get; }
        ICell this[string id] { get; }
        bool IsEmpty { get; }
        int Index { get; }
    }
}
