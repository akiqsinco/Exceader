using System;
using System.Collections.Generic;
using System.Xml;
using Exceader.Common;

namespace Exceader.Models
{
    public class Cell : ICell
    {
        public IRow Row { get; }

        public string Value { get; }

        public int Index { get; }

        public bool IsEmpty => string.IsNullOrEmpty(Value);

        public string Id => ExcelNumber.ToId(Row.Index, Index);

        internal Cell(IRow row, int index)
        {
            Row = row;
            Value = string.Empty;
            Index = index;
        }

        internal Cell(IRow row, int index, XmlElement cellElement, IReadOnlyList<string> sharedStrings)
        {
            Row = row;
            Value = GetCellValue(cellElement, sharedStrings);
            Index = index;
        }

        private string GetCellValue(XmlElement cellElement, IReadOnlyList<string> sharedStrings)
        {
            var value = cellElement["v"]?.InnerText;
            if (value == null)
            {
                return string.Empty;
            }

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
