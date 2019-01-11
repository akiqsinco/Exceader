using System;
using System.Text;

namespace Exceader.Models
{
    public struct CellPosition
    {
        #region static

        public static CellPosition Parse(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                throw new ArgumentNullException(nameof(id));
            }

            return TryParse(id, out var position) ? position : throw new FormatException();
        }

        public static bool TryParse(string id, out CellPosition position)
        {
            if (string.IsNullOrEmpty(id))
            {
                throw new ArgumentNullException(nameof(id));
            }

            var row = 0;
            var column = 0;
            var success = true;
            for (var i = 0; i < id.Length; i++)
            {
                var c = id[i];
                if ('A' <= c && c <= 'Z')
                {
                    column = column * 26 + (c - 'A' + 1);
                    continue;
                }

                success = i > 0 && int.TryParse(id.Substring(i), out row);
                break;
            }
            success &= row > 0;

            position = success ? new CellPosition(row - 1, column - 1) : default(CellPosition);
            return success;
        }

        #endregion

        private readonly int? _row;

        private readonly int? _column;

        public int Row => _row ?? int.MinValue;

        public int Column => _column ?? int.MinValue;

        public CellPosition(int row, int column)
        {
            if (row < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(row));
            }

            if (column < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            _row = row;
            _column = column;
        }

        public string ToId()
        {
            if (!_row.HasValue || !_column.HasValue)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            var n = _column.Value + 1;
            while (n > 0)
            {
                sb.Insert(0, (char)((n - 1) % 26 + 'A'));
                n = (n - 1) / 26;
            }
            sb.Append(_row.Value + 1);

            return sb.ToString();
        }

        public override string ToString() => ToId();
    }
}
