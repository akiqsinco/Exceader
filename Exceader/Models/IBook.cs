using System;
using System.Collections.Generic;

namespace Exceader.Models
{
    public interface IBook : IDisposable
    {
        ISheet this[string sheetName] { get; }
        IEnumerable<string> SheetNames { get; }
    }
}
