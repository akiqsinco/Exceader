using Exceader.Models;
using System;
using System.Globalization;

namespace Exceader.Extensions
{
    public static class CellExtensions
    {
        public static int? AsInteger(this ICell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            return int.TryParse(cell.Value, out var value) ? value : default(int?);
        }

        public static float? AsFloat(this ICell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            return float.TryParse(cell.Value, out var value) ? value : default(float?);
        }

        public static double? AsDouble(this ICell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            return double.TryParse(cell.Value, out var value) ? value : default(double?);
        }

        public static DateTime? AsDateTime(this ICell cell, IFormatProvider provider = null)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            return DateTime.TryParse(cell.Value, provider, DateTimeStyles.None, out var value) ? value : default(DateTime?);
        }

        public static DateTime? AsDateTime(this ICell cell, string format, IFormatProvider provider = null)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            return DateTime.TryParseExact(cell.Value, format, provider, DateTimeStyles.None, out var value) ? value : default(DateTime?);
        }
    }
}
