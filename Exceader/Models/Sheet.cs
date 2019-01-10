using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Models
{
    public class Sheet : ISheet
    {
        public static readonly ISheet Empty = new Sheet();

        private readonly IRow[] _rows;

        public IRow this[int index] => throw new NotImplementedException();

        public ICell this[string id] => throw new NotImplementedException();

        public IBook Book => throw new NotImplementedException();

        public bool IsEmpty => _rows.All(row => row.IsEmpty);

        public int Index => throw new NotImplementedException();

        public int Count => _rows.Length;

        public IEnumerator<IRow> GetEnumerator()
        {
            return ((IEnumerable<IRow>)_rows).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rows.GetEnumerator();
        }
    }
}
