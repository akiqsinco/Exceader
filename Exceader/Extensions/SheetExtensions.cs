using Exceader.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Exceader.Extensions
{
    public static class SheetExtensions
    {
        public static IEnumerable<IRow> Range(this ISheet sheet, int from, int to)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (to < from)
            {
                throw new InvalidOperationException();
            }

            return Enumerable.Range(from, to - from).Select(i => sheet[i]);
        }
    }
}
