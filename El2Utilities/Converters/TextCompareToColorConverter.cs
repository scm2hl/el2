using El2Core.Models;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace El2Core.Converters
{
    [ValueConversion(typeof(string[]), typeof(SolidColorBrush))]
    public class TextCompareToColorConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {

            if (values == null || values.Length != 2) return false;
            Vorgang v1 = (Vorgang)values[0];
            string v2 = (string)values[1];
            if (!string.IsNullOrEmpty(v2) && (v1.Aid.Contains(v2, StringComparison.CurrentCultureIgnoreCase)
                || (v1.AidNavigation?.Material?.Contains(v2, StringComparison.CurrentCultureIgnoreCase) ?? false)
                || (v1.AidNavigation?.MaterialNavigation?.Bezeichng?.Contains(v2, StringComparison.CurrentCultureIgnoreCase) ?? false)))
            {
                return Brushes.Khaki;
            }
            else return null;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
