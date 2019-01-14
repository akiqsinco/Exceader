using System;
using System.Collections.Generic;
using System.Xml;

namespace Exceader.Models
{
    public class Cell : ICell
    {
        private CellPosition _position;

        public IRow Row { get; }

        public string Value { get; }

        public int Index => _position.Column;

        public string Id => _position.ToId();

        public bool IsEmpty => string.IsNullOrEmpty(Value);

        internal Cell(IRow row, int index)
        {
            Row = row;
            Value = string.Empty;
            _position = new CellPosition(row.Index, index);
        }

        internal Cell(IRow row, int index, XmlElement cellElement, IReadOnlyList<string> sharedStrings)
        {
            Row = row;
            Value = GetCellValue(cellElement, sharedStrings);
            _position = new CellPosition(row.Index, index);
        }

        private string GetCellValue(XmlElement cellElement, IReadOnlyList<string> sharedStrings)
        {
            var value = cellElement["v"].InnerText;
            var type = cellElement.GetAttribute("t");
            if (type == null || type != "s")
            {
                return value;
            }

            if (!int.TryParse(value, out var refIndex))
            {
                return string.Empty;
            }

            if (refIndex < 0 || sharedStrings.Count <= refIndex)
            {
                return string.Empty;
            }

            return sharedStrings[refIndex];
        }

        public override string ToString() => Value;
    }
}
