using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace El2Core.Utils
{
    public static class Extensions
    {
        /// <summary>
        /// Compares two objects of the same type and returns a list of variances, which includes the property name and the differing values.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="val1"></param>
        /// <param name="val2"></param>
        /// <returns></returns>
        public static List<Variance> DetailedCompare<T>(this T val1, T val2)
        {
            List<Variance> variances = new List<Variance>();
            FieldInfo[] fi = val1.GetType().GetFields();
            foreach (FieldInfo f in fi)
            {
                Variance v = new Variance();
                v.Prop = f.Name;
                v.valA = f.GetValue(val1);
                v.valB = f.GetValue(val2);
                if (!Equals(v.valA, v.valB))
                    variances.Add(v);

            }
            return variances;
        }
        /// <summary>
        /// Compares two JSON objects represented as strings and returns a new JsonElement that represents the differences between them.
        /// </summary>
        /// <param name="json1"></param>
        /// <param name="json2"></param>
        /// <returns></returns>
        public static JsonElement CompareJsonObjects(string json1, string json2)
        {
            using var doc1 = JsonDocument.Parse(json1);
            using var doc2 = JsonDocument.Parse(json2);

            var obj1 = doc1.RootElement;
            var obj2 = doc2.RootElement;

            return CompareJsonElements(obj1, obj2);
        }
        /// <summary>
        /// Compares two JsonElements and returns a new JsonElement that represents the differences between them.
        /// </summary>
        /// <param name="elem1"></param>
        /// <param name="elem2"></param>
        /// <returns></returns>
        private static JsonElement CompareJsonElements(JsonElement elem1, JsonElement elem2)
        {
            if (elem1.ValueKind != elem2.ValueKind)
                return CreateDifferenceArray(elem1, elem2);

            switch (elem1.ValueKind)
            {
                case JsonValueKind.Object:
                    var result = new Dictionary<string, JsonElement>();
                    foreach (var prop in elem1.EnumerateObject())
                        if (elem2.TryGetProperty(prop.Name, out var value2))
                        {
                            var diff = CompareJsonElements(prop.Value, value2);
                            if (diff.ValueKind != JsonValueKind.Undefined)
                                result.Add(prop.Name, diff);
                        }
                        else
                            result.Add(prop.Name, CreateDifferenceArray(prop.Value, default));

                    foreach (var prop in elem2.EnumerateObject().Where(prop => !elem1.TryGetProperty(prop.Name, out _)))
                        result.Add(prop.Name, CreateDifferenceArray(default, prop.Value));

                    return JsonElementFromObject(result);

                case JsonValueKind.Array:
                    var areArraysEqual = elem1.GetArrayLength() == elem2.GetArrayLength();
                    if (areArraysEqual)
                        for (var i = 0; i < elem1.GetArrayLength(); i++)
                            if (!elem1[i].Equals(elem2[i]))
                            {
                                areArraysEqual = false;
                                break;
                            }

                    return areArraysEqual ? default : CreateDifferenceArray(elem1, elem2);

                default:
                    return elem1.Equals(elem2) ? default : CreateDifferenceArray(elem1, elem2);
            }
        }
        /// <summary>
        /// Creates a JsonElement that represents the difference between two JsonElements by placing them in an array.
        /// </summary>
        /// <param name="elem1"></param>
        /// <param name="elem2"></param>
        /// <returns></returns>
        private static JsonElement CreateDifferenceArray(JsonElement elem1, JsonElement elem2)
        {
            var array = new JsonElement[] { elem1, elem2 };
            return JsonElementFromObject(array);
        }
        /// <summary>
        /// Converts an object to a JsonElement by serializing it to a JSON string and then parsing that string back into a JsonElement.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private static JsonElement JsonElementFromObject(object obj)
        {
            var jsonString = JsonSerializer.Serialize(obj);
            using var doc = JsonDocument.Parse(jsonString);
            return doc.RootElement.Clone();
        }

    }
    public class Variance
    {
        public string Prop { get; set; }
        public object valA { get; set; }
        public object valB { get; set; }
    }
}

