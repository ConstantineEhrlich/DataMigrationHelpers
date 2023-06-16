using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Serialization
{
    /// <summary>
    /// This class provides methods to serialize objects to dictionary or array.
    /// </summary>
    public static class Serializer
    {
        /// <summary>
        /// Gets a dictionary where the keys are property names of the provided type and the values are their respective indices.
        /// Only properties of serializable types are considered.
        /// </summary>
        /// <param name="type">The type of the object to get the property map for.</param>
        /// <returns>A dictionary mapping property names to their indices.</returns>
        public static Dictionary<string,int> GetPropertyMap(Type type)
        {
            List<string> serializableTypes = new() { "DateTime", "String", "Int16", "Int32", "Int64", "Single", "Double", "Decimal" };

            string[] properties = type.GetProperties()
                                      .Where(prop => prop.CanRead)
                                      .Where(prop => serializableTypes.Contains(Nullable.GetUnderlyingType(prop.PropertyType)?.Name ?? prop.PropertyType.Name))
                                      .Select(prop => prop.Name)
                                      .ToArray();
            return properties.Select((prop, index) => new { k = prop, v = index }).ToDictionary(kv => kv.k, kv => kv.v);
        }

        /// <summary>
        /// Serializes an object of the given type to an array.
        /// </summary>
        /// <typeparam name="T">The type of the object to be serialized.</typeparam>
        /// <param name="obj">The object to serialize.</param>
        /// <returns>An array of object values representing the serialized object.</returns>
        public static object?[] SerializeToArray<T>(T obj)
        {
            Dictionary<string, int> map = GetPropertyMap(typeof(T));
            object?[] result = new object?[map.Count];
            int i = 0;
            foreach (string propName in map.Keys)
            {
                PropertyInfo prop = typeof(T).GetProperty(propName)!;
                result[i] = prop.GetValue(obj);
                i++;
            }

            return result;
        }


        /// <summary>
        /// Serializes an object of the given type to a dictionary.
        /// </summary>
        /// <typeparam name="T">The type of the object to be serialized.</typeparam>
        /// <param name="obj">The object to serialize.</param>
        /// <returns>A dictionary where the keys are the property names and the values are their respective values from the serialized object.</returns>
        public static Dictionary<string, object?> SerializeToDict<T>(T obj)
        {
            Dictionary<string, object?> result = new();
            Dictionary<string, int> map = GetPropertyMap(typeof(T));
            object?[] vals = SerializeToArray<T>(obj);
            int i = 0;
            foreach (string prop in map.Keys)
            {
                result.Add(prop, vals[i]);
                i++;                    
            }
            return result;
        }

    }
}
