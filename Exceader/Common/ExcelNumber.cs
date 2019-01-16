using System;
using System.Text;

namespace Exceader.Common
{
    public static class ExcelNumber
    {
        public static int ToNumerical(string alphabeticalIndex)
        {
            if (alphabeticalIndex == null)
            {
                throw new ArgumentNullException(nameof(alphabeticalIndex));
            }

            var numerical = 0;
            var success = true;
            foreach (var c in alphabeticalIndex)
            {
                if ('A' <= c && c <= 'Z')
                {
                    numerical = numerical * 26 + (c - 'A' + 1);
                }
                else
                {
                    success = false;
                    break;
                }
            }
            success &= numerical > 0;

            if (!success)
            {
                throw new FormatException();
            }

            return numerical - 1;
        }

        public static string ToAlphabetical(int numericalIndex)
        {
            if (numericalIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(numericalIndex));
            }

            var sb = new StringBuilder();
            var n = numericalIndex + 1;
            while (n > 0)
            {
                sb.Insert(0, (char)((n - 1) % 26 + 'A'));
                n = (n - 1) / 26;
            }

            return sb.ToString();
        }

        public static (int row, int column) ToIndexes(string cellId)
        {
            if (cellId == null)
            {
                throw new ArgumentNullException(nameof(cellId));
            }

            var row = 0;
            var column = 0;
            var success = true;
            for (var i = 0; i < cellId.Length; i++)
            {
                var c = cellId[i];
                if ('A' <= c && c <= 'Z')
                {
                    column = column * 26 + (c - 'A' + 1);
                    continue;
                }

                success = i > 0 && int.TryParse(cellId.Substring(i), out row);
                break;
            }
            success &= row > 0;

            if (!success)
            {
                throw new FormatException();
            }

            return (row - 1, column - 1);
        }

        public static string ToCellId(int row, int column)
        {
            if (row < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(row));
            }

            if (column < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            return $"{ToAlphabetical(column)}{row + 1}";
        }
    }
}
