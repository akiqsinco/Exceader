using System;
using System.Collections.Generic;
using System.Xml;
using System.Linq;
using Exceader.Common;

namespace Exceader.Models
{
    public class Row : IRow
    {
        private readonly IReadOnlyDictionary<int, ICell> _cells;

        public ICell this[int index]
        {
            get
            {
                if (index < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                if (_cells.ContainsKey(index))
                {
                    return _cells[index];
                }

                return new Cell(this, index);
            }
        }

        public ISheet Sheet { get; }

        public int Index { get; }

        public bool IsEmpty => _cells.Values.All(cell => cell.IsEmpty);

        internal Row(ISheet sheet, int index)
        {
            Sheet = sheet;
            Index = index;
            _cells = new Dictionary<int, ICell>();
        }

        internal Row(ISheet sheet, int index, XmlElement rowElement, IReadOnlyList<string> sharedStrings)
        {
            Sheet = sheet;
            Index = index;
            _cells = CreateCells(rowElement, sharedStrings);
        }

        private Dictionary<int, ICell> CreateCells(XmlElement rowElement, IReadOnlyList<string> sharedStrings)
        {
            var cells = new Dictionary<int, ICell>();

            foreach (XmlElement c in rowElement.GetElementsByTagName("c"))
            {
                var (_, column) = ExcelNumber.ToIndexes(c.GetAttribute("r"));

                cells[column] = new Cell(this, column, c, sharedStrings);
            }

            return cells;
        }
    }
}
