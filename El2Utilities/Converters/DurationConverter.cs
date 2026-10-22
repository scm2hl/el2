using System;
using System.Globalization;
using System.Windows.Data;

namespace El2Core.Converters
{
    public class DurationConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            Double d = 0.0;
            Double corr = 0.0;
            if (values == null || values.Length < 4)
                return "NaN";
            if (values[1] == null || !int.TryParse(values[1]?.ToString(), out var rstze))
                return "NaN";
            if (values[2] != null && !Double.TryParse(values[2]?.ToString(), out corr))
                return "NaN";
            if (values[3] != null && Single.TryParse(values[3]?.ToString(), out var quant))
            {
                
                d = (Single)values[0] + corr * quant + rstze;

            } else return "NaN";

            return string.Format("{0:F2}h", d / 60);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
