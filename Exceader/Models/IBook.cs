using System;
using System.Collections.Generic;

namespace Exceader.Models
{
    public interface IBook : IReadOnlyList<ISheet>, IDisposable
    {
        ISheet this[string sheetName] { get; }
    }
}
