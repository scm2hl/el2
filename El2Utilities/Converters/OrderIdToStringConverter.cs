using System;
using System.Globalization;
using System.Windows.Data;

namespace El2Core.Converters
{
    public class OrderIdToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return string.Empty;

            if (value is long l)
                return l == 0 ? string.Empty : l.ToString(culture);

            if (value is int i)
                return i == 0 ? string.Empty : i.ToString(culture);

            return value.ToString() ?? string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = (value as string)?.Trim();
            if (string.IsNullOrEmpty(s))
            {
                // return 0 for empty input
                if (targetType == typeof(long) || targetType == typeof(long?))
                    return 0L;
                if (targetType == typeof(int) || targetType == typeof(int?))
                    return 0;
                return 0L;
            }

            if (long.TryParse(s, NumberStyles.Integer, culture, out var l))
            {
                if (targetType == typeof(int) && l <= int.MaxValue)
                    return (int)l;
                return l;
            }

            // fallback to 0 on parse failure
            if (targetType == typeof(int) || targetType == typeof(int?))
                return 0;
            return 0L;
        }
    }
}
