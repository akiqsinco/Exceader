using System;
using System.Collections.Generic;
using System.Xml;
using System.Linq;
using Exceader.Common;

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

                var (row, column) = ExcelNumber.ToIndexes(id);

                return this[row][column];
            }
        }

        public ICell this[int row, int column] => this[row][column];

        public IBook Book { get; }

        public bool IsEmpty => _rows.Values.All(row => row.IsEmpty);

        public string Name { get; }

        internal Sheet(IBook book, string name, XmlElement sheetElement, IReadOnlyList<string> sharedStrings)
        {
            Book = book;
            Name = name;
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
