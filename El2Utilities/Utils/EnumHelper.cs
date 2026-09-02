using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace El2Core.Utils
{
    /// <summary>
    /// Provides helper methods for working with enums, including retrieving descriptions and filtering based on attributes.
    /// </summary>
    public static class EnumHelper
    {
        // Cache must consider the descriptionParameter because callers may request different
        // filtering (only browsable entries vs. all entries). Use a composite string key.
        private static readonly ConcurrentDictionary<string, List<ValueDescription>> _cache = new();

        public static string Description(this Enum value)
        {
            if (value == null) return string.Empty;

            var field = value.GetType().GetField(value.ToString());
            var attr = field?.GetCustomAttribute<DescriptionAttribute>(false);
            if (!string.IsNullOrEmpty(attr?.Description))
                return attr.Description!;

            // Fallback: make a readable name from the enum identifier
            var ti = CultureInfo.CurrentCulture.TextInfo;
            return ti.ToTitleCase(ti.ToLower(value.ToString().Replace("_", " ")));
        }

        // Keep signature compatible with existing callers. The parameter is currently unused
        // but preserved to avoid breaking XAML converter usages that pass a parameter.
        public static IEnumerable<ValueDescription> GetAllValuesAndDescriptions(Type t, string? descriptionParameter)
        {
            if (!t.IsEnum)
                throw new ArgumentException($"{nameof(t)} must be an enum type");
            var key = $"{t.FullName}|{descriptionParameter ?? string.Empty}";

            return _cache.GetOrAdd(key, _ =>
            {
                var fields = t.GetFields(BindingFlags.Public | BindingFlags.Static);

                IEnumerable<FieldInfo> filtered;
                // If parameter equals "1" only include members explicitly marked Browsable(true).
                if (descriptionParameter == "1")
                {
                    filtered = fields.Where(f =>
                    {
                        var b = f.GetCustomAttribute<BrowsableAttribute>(false);
                        return b != null && b.Browsable;
                    });
                }
                else
                {
                    // Include all enum fields
                    filtered = fields;
                }

                return filtered
                    .Select(f =>
                    {
                        var val = (Enum)f.GetValue(null)!;
                        return new ValueDescription { Value = val, Description = val.Description() };
                    })
                    .ToList();
            });
        }

        public static IEnumerable<ValueDescription> GetAllValuesAndDescriptions<TEnum>(string? descriptionParameter = null) where TEnum : Enum
            => GetAllValuesAndDescriptions(typeof(TEnum), descriptionParameter);
    }
}
