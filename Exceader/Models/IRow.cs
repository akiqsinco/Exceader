using System.Collections.Generic;

namespace Exceader.Models
{
    public interface IRow
    {
        ISheet Sheet { get; }
        ICell this[int index] { get; }
    }
}
