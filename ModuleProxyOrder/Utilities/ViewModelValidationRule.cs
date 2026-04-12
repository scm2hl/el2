using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ModuleProxyOrder.Utilities
{
    public class ViewModelValidationRule : ValidationRule
    {
        // Name of the view element (UserControl) to find so we can get its DataContext (the ViewModel)
        public string ViewElementName { get; set; } = string.Empty;

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            try
            {
                if (string.IsNullOrEmpty(ViewElementName))
                    return ValidationResult.ValidResult;

                // Find the element by name in the visual tree of open windows
                var view = Application.Current.Windows
                    .OfType<Window>()
                    .SelectMany(w => w.FindName(ViewElementName) is FrameworkElement fe ? new[] { fe } : Array.Empty<FrameworkElement>())
                    .FirstOrDefault();

                if (view == null)
                {
                    // try to find non-window elements
                    var fe2 = Application.Current.Windows
                        .OfType<Window>()
                        .Select(w => w.Content)
                        .OfType<FrameworkElement>()
                        .SelectMany(root => FindByNameRecursive(root, ViewElementName))
                        .FirstOrDefault();

                    view = fe2;
                }

                if (view == null)
                {
                    return ValidationResult.ValidResult;
                }

                if (view.DataContext is System.Windows.Controls.ValidationRule vmRule)
                {
                    // If the view model itself inherits ValidationRule (like ViewModelBase does), call it
                    return vmRule.Validate(value, cultureInfo);
                }

                // If DataContext has a Validate method (like in ProxyOrderViewModel), try to invoke it
                var dc = view.DataContext;
                if (dc == null)
                    return ValidationResult.ValidResult;

                var method = dc.GetType().GetMethod("Validate", new Type[] { typeof(object), typeof(CultureInfo) });
                if (method != null)
                {
                    var result = method.Invoke(dc, new object[] { value, cultureInfo }) as ValidationResult;
                    if (result != null)
                        return result;
                }

                return ValidationResult.ValidResult;
            }
            catch (Exception ex)
            {
                return new ValidationResult(false, ex.Message);
            }
        }

        private static System.Collections.Generic.IEnumerable<FrameworkElement> FindByNameRecursive(FrameworkElement root, string name)
        {
            if (root == null) yield break;
            if (root.Name == name) yield return root;
            int count = VisualTreeHelper.GetChildrenCount(root);
            for (int i = 0; i < count; i++)
            {
                if (VisualTreeHelper.GetChild(root, i) is FrameworkElement child)
                {
                    foreach (var found in FindByNameRecursive(child, name))
                        yield return found;
                }
            }
        }
    }
}
