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

            if (from < to)
            {
                throw new InvalidOperationException();
            }

            return Enumerable.Range(from, to - from).Select(i => row[i]);
        }
    }
}
