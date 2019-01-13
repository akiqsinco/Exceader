using System;
using System.Collections.Generic;
using System.Xml;

namespace Exceader.Models
{
    public class Sheet : ISheet
    {
        private readonly IReadOnlyDictionary<int, IRow> _rows;

        public IRow this[int index]
        {
            get
            {
                if (index < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                if (_rows.ContainsKey(index))
                {
                    return _rows[index];
                }

                return new Row(this, index);
            }
        }

        public ICell this[string id]
        {
            get
            {
                if (id == null)
                {
                    throw new ArgumentNullException(nameof(id));
                }

                if (!CellPosition.TryParse(id, out var pos))
                {
                    throw new FormatException();
                }

                return this[pos.Row][pos.Column];
            }
        }

        public ICell this[int row, int column] => this[row, column];

        public IBook Book { get; }

        internal Sheet(IBook book, XmlElement sheetElement, IReadOnlyList<string> sharedStrings)
        {
            Book = book;
            _rows = CreateRows(sheetElement, sharedStrings);
        }

        private Dictionary<int, IRow> CreateRows(XmlElement sheetElement, IReadOnlyList<string> sharedStrings)
        {
            var rows = new Dictionary<int, IRow>();

            foreach (XmlElement row in sheetElement.GetElementsByTagName("row"))
            {
                var index = int.Parse(row.GetAttribute("r")) - 1;

                rows[index] = new Row(this, index, row, sharedStrings);
            }

            return rows;
        }
    }
}
