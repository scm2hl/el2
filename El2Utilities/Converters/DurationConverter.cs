using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Data;

namespace El2Core.Converters
{
    public class DurationConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            Double d = 0.0;
            if (values == null || values.Length < 3)
                return "NaN";
            if (Single.TryParse(values[1] as string, out var basemg))
            {
                
                d = (Single)values[0] / basemg * (int)values[2];
                foreach (var item in values.Skip(3))
                {

                    if (item != null)
                        if (item is Single s)
                        {
                            d += s;
                        }
                
                }
            } else return "NaN";

            return string.Format("{0:F2}h", d / 60);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
