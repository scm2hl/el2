using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;

namespace El2Core.Utils
{
    public static class EnumHelper
    {
        public static string Description(this Enum value)
        {
            if (value == null) return string.Empty;

            var field = value.GetType().GetField(value.ToString());
            var attr = field?.GetCustomAttribute<DescriptionAttribute>(false);
            if (attr != null && !string.IsNullOrEmpty(attr.Description))
                return attr.Description;

            // Fallback: make a readable name from the enum identifier
            TextInfo ti = CultureInfo.CurrentCulture.TextInfo;
            return ti.ToTitleCase(ti.ToLower(value.ToString().Replace("_", " ")));
        }

        public static IEnumerable<ValueDescription> GetAllValuesAndDescriptions(Type t, string? descriptionParameter)
        {
            if (!t.IsEnum)
                throw new ArgumentException($"{nameof(t)} must be an enum type");

            // Return enum values where the member is either not decorated with BrowsableAttribute
            // or the BrowsableAttribute.Browsable == true. If a member has [Browsable(false)] it will be skipped.
            var fields = t.GetFields(BindingFlags.Public | BindingFlags.Static);
            var list = new List<ValueDescription>();
            var des = descriptionParameter ?? string.Empty;
            foreach (var field in fields)
            {
                var browsable = field.GetCustomAttribute<EditorBrowsableAttribute>();
                if (browsable != null && des.Equals("1"))
                    continue; // skip non-browsable members

                var value = (Enum)field.GetValue(null);
                list.Add(new ValueDescription { Value = value, Description = value.Description() });
            }

            return list;
        }
    }
}
