using System.Collections.Generic;

namespace Exceader.Models
{
    public interface IRow : IReadOnlyList<ICell>
    {
        ISheet Sheet { get; }
        bool IsEmpty { get; }
        int Index { get; }
    }
}
