using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Reflection;
using System.Text.Json;

using ExcelHelpers;

namespace Serialization
{
    /// <summary>
    /// This class provides methods to deserialize data record objects.
    /// </summary>
    public static class Deserializer
    {
        /// <summary>
        /// Deserializes a data record into an object of the specified type.
        /// </summary>
        /// <typeparam name="T">The type of the object to be deserialized into.</typeparam>
        /// <param name="reader">The data record to deserialize.</param>
        /// <param name="headerMap">Specifies how the header map should be obtained.</param>
        /// <param name="jsonMap">Specifies the json map, if required.</param>
        /// <returns>An object of the specified type, populated with the data from the data record.</returns>
        /// <exception cref="NullReferenceException">Thrown if a matching constructor for the target type can't be found.</exception>
        public static T Deserialize<T>(IDataRecord reader,
                                       HeaderMap headerMap = HeaderMap.ColumnHeaders,
                                       string jsonMap = "")
        {
            Type targetType = typeof(T);

            // Get the map
            Dictionary<string, int> readerMap = headerMap switch
            {
                HeaderMap.ColumnHeaders => GetMapFromReader(reader),
                HeaderMap.JsonMap => DataReaderFactory.GetMapFromJson(jsonMap, reader.FieldCount),
                _ => new()
            };
            
            Dictionary<string, int> propertiesMap = FilterByTypeProperties(readerMap, targetType);
                        
            // Create object of type
            ConstructorInfo constructor = targetType.GetConstructor(Array.Empty<Type>())
                ?? throw new NullReferenceException($"Can't find a constructor for type {typeof(T).Name} that takes 0 arguments !");

            T resultObject = (T)constructor.Invoke(null);

            // Assign properties
            foreach ((string propName, int i) in propertiesMap)
            {
                PropertyInfo prop = targetType.GetProperty(propName)!;
                if (reader.IsDBNull(propertiesMap[propName]))
                {
                    prop.SetValue(resultObject, null);
                    continue;
                }

                // Since some properties are nullable, we can't get the type directly from the property
                // prop.PropertyType returns "Nullable`1". Therefore, this is the solution to get the underlying type:
                Type type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                object? value;

                if (type.Name == "Decimal")
                {
                    value = reader.GetDataTypeName(i) switch
                    {
                        "int" => Convert.ToDecimal(reader.GetInt32(i)),
                        "real" => Convert.ToDecimal(Math.Round(reader.GetFloat(i), 6)),
                        "float" => Convert.ToDecimal(Math.Round(reader.GetDouble(i), 6)),
                        "decimal" => reader.GetDecimal(i),
                        _ => reader.GetDecimal(i),
                    };
                }
                else
                {
                    value = type.Name switch
                    {
                        "DateTime" => reader.GetDateTime(i),
                        "String" => reader.GetString(i),
                        "Int16" => reader.GetInt16(i),
                        "Int32" => reader.GetInt32(i),
                        "Int64" => reader.GetInt64(i),
                        "Single" => reader.GetFloat(i),
                        "Double" => reader.GetDouble(i),
                        _ => null
                    };
                }

                

                prop.SetValue(resultObject, value);
            }

            return resultObject;

        }



        /// <summary>
        /// Generates a map between the properties of the object and their indices in the data record.
        /// </summary>
        /// <param name="reader">The data record to create a map from.</param>
        /// <returns>A dictionary where keys are property names and values are their respective indices.</returns>
        private static Dictionary<string, int> GetMapFromReader(IDataRecord reader)
        {
            if (reader.GetType() == typeof(DataReaderFactory))
            {
                DataReaderFactory drf = (DataReaderFactory)reader;
                return drf.DictMap!;
            }
            else
            {
                Dictionary<string, int> map = new();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    map.Add(reader.GetName(i), i);
                }

                return map;
            }
        }


        /// <summary>
        /// Filters the provided map by properties of the specified type that can be set.
        /// </summary>
        /// <param name="readerMap">The map to filter.</param>
        /// <param name="type">The type of the object to filter the map for.</param>
        /// <returns>A dictionary where keys are property names and values are their respective indices.</returns>
        private static Dictionary<string, int> FilterByTypeProperties(Dictionary<string, int> readerMap, Type type)
        {
            string[] properties = type.GetProperties()
                                      .Where(prop => prop.CanWrite)
                                      //.Where(prop => !prop.PropertyType.IsGenericType)
                                      .Select(prop => prop.Name)
                                      .ToArray();

            return readerMap.Where(kv => properties.Contains(kv.Key))
                            .ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        
    }
}
