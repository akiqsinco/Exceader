using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Models
{
    public class Row : IRow
    {
        public static readonly IRow Empty = new Row();

        private readonly ICell[] _cells;

        public ICell this[int index] => throw new NotImplementedException();

        public ISheet Sheet => throw new NotImplementedException();

        public bool IsEmpty => _cells.All(cell => cell.IsEmpty);

        public int Index => throw new NotImplementedException();

        public int Count => _cells.Length;

        public IEnumerator<ICell> GetEnumerator()
        {
            return ((IEnumerable<ICell>)_cells).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _cells.GetEnumerator();
        }
    }
}
