﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Serialization
{
    public static class Serializer
    {
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
