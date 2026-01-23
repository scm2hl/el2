using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Data;

namespace El2Core.Converters
{
    public partial class ReferenceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string s)

            {
                
                var ss = s.Split((char)29);
                Regex reg = new Regex("w+");
                var r = Regex.Match(ss[2], @"^([\w\-]+)").Value;
                return r;
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

  
    }
}
