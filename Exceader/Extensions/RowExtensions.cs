using Exceader.Common;
using Exceader.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Exceader.Extensions
{
    public static class RowExtensions
    {
        public static IEnumerable<ICell> Range(this IRow row, int from, int to)
        {
            if (row == null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            if (to < from)
            {
                throw new InvalidOperationException($"'{nameof(from)}' should be less or equal to '{nameof(to)}'");
            }

            return Enumerable.Range(from, to - from + 1).Select(i => row[i]);
        }

        public static IEnumerable<ICell> Range(this IRow row, string from, string to)
        {
            if (row == null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            if (from == null)
            {
                throw new ArgumentNullException(nameof(from));
            }

            if (to == null)
            {
                throw new ArgumentNullException(nameof(to));
            }

            var start = ExcelNumber.ToNumerical(from);
            var end = ExcelNumber.ToNumerical(to);
            if (end < start)
            {
                throw new InvalidOperationException($"'{nameof(from)}' should be less or equal to '{nameof(to)}'");
            }

            return Enumerable.Range(start, end - start + 1).Select(i => row[i]);
        }
    }
}
