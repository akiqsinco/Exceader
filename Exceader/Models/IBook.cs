using System;
using System.Collections.Generic;

namespace Exceader.Models
{
    public interface IBook : IDisposable
    {
        ISheet this[int index] { get; }
        ISheet this[string sheetName] { get; }
    }
}
