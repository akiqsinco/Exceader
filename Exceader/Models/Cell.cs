using System;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Models
{
    public class Cell : ICell
    {
        public static readonly ICell Empty = new Cell();

        public IRow Row => throw new NotImplementedException();

        public string Value => throw new NotImplementedException();

        public bool IsEmpty => string.IsNullOrEmpty(Value);

        public int Index => throw new NotImplementedException();

        public string Id => throw new NotImplementedException();
    }
}
