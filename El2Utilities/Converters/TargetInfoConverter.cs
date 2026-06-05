using El2Core.Models;
using System;
using System.Globalization;
using System.Windows.Data;

namespace El2Core.Converters
{
    public class TargetInfoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is EmployeeNote em)
                return em.GetTargetInfo() ?? string.Empty;
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
