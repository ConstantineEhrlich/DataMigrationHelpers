using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using System.Text.Json;

namespace ExcelHelpers
{
    /// <summary>
    /// Provides an implementation of the <see cref="IDataReader"/> interface that reads data from in-memory collections.
    /// </summary>
    /// <remarks>
    /// This class is used to provide data from various collection types, including lists of arrays and dictionaries,
    /// and to expose it in a form of a data reader.
    /// </remarks>
    /// <exception cref="NotSupportedException">Thrown when trying to access unsupported methods like
    ///                                         GetData, GetBoolean, GetByte, GetBytes, GetChar, GetChars, GetGuid.</exception>
    /// <exception cref="NullReferenceException">Thrown when trying to access JsonMap or DictMap without setting it.</exception>
    /// <exception cref="FormatException">Thrown when provided JsonMap is not valid.</exception>
    public class DataReaderFactory : IDataReader
    {
        #region Fields

        private Dictionary<string, int>? _map;
        private IEnumerable<object?[]> _records;
        private IEnumerator<object?[]>? _recordsEnumerator;

        #endregion

        
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the DataReaderFactory class with the given records.
        /// </summary>
        /// <param name="records">Collection of object arrays to be read by the DataReaderFactory instance.</param>
        public DataReaderFactory(IEnumerable<object?[]> records)
        {
            _records = records;
        }

        
        /// <summary>
        /// Initializes a new instance of the DataReaderFactory class with the given records and HeaderSource.
        /// </summary>
        /// <param name="records">Collection of object arrays to be read by the DataReaderFactory instance.</param>
        /// <param name="headerSource">The <see cref="HeaderSource"/> source of the header information for the records.</param>
        public DataReaderFactory(IEnumerable<object?[]> records, HeaderSource headerSource)
        {
            _records = records;
            HeaderSource = headerSource;
        }

        
        /// <summary>
        /// Initializes a new instance of the DataReaderFactory class with the given records, HeaderSource, and jsonMap.
        /// </summary>
        /// <param name="records">The records to be read by the DataReaderFactory instance.</param>
        /// <param name="headerSource">The source of the header information for the records.</param>
        /// <param name="jsonMap">A JSON representation of a dictionary mapping from header names to column indices.</param>
        public DataReaderFactory(IEnumerable<object?[]> records, HeaderSource headerSource, string jsonMap)
        {
            _records = records;
            HeaderSource = headerSource;
            JsonMap = jsonMap;
        }

        
        /// <summary>
        /// Initializes a new instance of the DataReaderFactory class with the given records, HeaderSource, and map.
        /// </summary>
        /// <param name="records">The records to be read by the DataReaderFactory instance.</param>
        /// <param name="headerSource">The source of the header information for the records.</param>
        /// <param name="map">A dictionary mapping from header names to column indices.</param>
        public DataReaderFactory(IEnumerable<object?[]> records, HeaderSource headerSource, Dictionary<string, int> map)
        {
            _records = records;
            HeaderSource = headerSource;
            DictMap = map;
        }

        
        /// <summary>
        /// Initializes a new instance of the DataReaderFactory class with the given records, setting the HeaderSource to DictMap.
        /// </summary>
        /// <param name="records">The records to be read by the DataReaderFactory instance. Each record is represented as a dictionary mapping from header names to values.</param>
        public DataReaderFactory(IEnumerable<Dictionary<string, object?>> records)
        {
            HeaderSource = HeaderSource.DictMap;
            DictMap = records.First().Select((kv, index) => new { k = kv.Key, v = index }).ToDictionary(kv => kv.k, kv => kv.v);
            _records = records.Select(dict => dict.Values.ToArray());
        }

        #endregion



        #region Properties

        /// <summary>
        /// Gets or sets the source of the headers.
        /// </summary>
        public HeaderSource HeaderSource { get; set; } = HeaderSource.Noheader;

        /// <summary>
        /// Gets or sets the JSON map.
        /// </summary>
        public string? JsonMap { get; set; }

        /// <summary>
        /// Gets or sets the dictionary map.
        /// </summary>
        public Dictionary<string, int>? DictMap { get; set;}

        public int Depth { get => 0; }

        public bool IsClosed { get; private set; } = false;

        public int RecordsAffected { get => -1; }

        public int FieldCount { get => _map?.Count ?? 0; }

        #endregion



        #region Indexers

        public object this[string name] => _recordsEnumerator!.Current[_map![name]] ?? System.DBNull.Value;

        public object this[int i] => _recordsEnumerator!.Current[i] ?? System.DBNull.Value;

        #endregion



        #region Public Methods

        /// <summary>
        /// Creates a map of string and int from a JSON string.
        /// </summary>
        /// <param name="jsonMap">A JSON string representing the map to be created.</param>
        /// <param name="fieldCount">The number of fields expected in the map.</param>
        /// <returns>A dictionary that represents the JSON string.</returns>
        public static Dictionary<string, int> GetMapFromJson(string jsonMap, int fieldCount)
        {
            Dictionary<string, int> map = new();

            JsonSerializerOptions jsonOpts = new() { AllowTrailingCommas = true };
            Dictionary<string, string> json = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonMap, jsonOpts)
                ?? throw new FormatException($"Invalid json");

            // Add to the map only values that are not causing exception in LetterToIndex method
            foreach (var kv in json)
            {
                try
                {
                    int colIndex = ExcelExtensions.LetterToIndex(kv.Value);
                    map.Add(kv.Key, colIndex);
                }
                catch (Exception)
                {
                    continue;
                }
            }

            // Trim from the map the values that are smaller than field count
            return map.Where(kv => kv.Value < fieldCount)
                      .ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        public bool Read()
        {
            if (_map is null)
            {
                _map = GetHeaderMap();
            }

            if (_recordsEnumerator is null)
            {
                _recordsEnumerator = _records.GetEnumerator();
            }

            return _recordsEnumerator.MoveNext();
        }

        public int GetValues(object[] values)
        {
            int i = 0;
            foreach (string field in _map!.Keys)
            {
                values[i] = this[field];
                i++;
            }
            return FieldCount;
        }

        public string GetName(int i) => _map!.ElementAt(i).Key;
        public int GetOrdinal(string name) => _map![name];

        public bool NextResult() => false;
        public void Close() => IsClosed = true;
        public DataTable? GetSchemaTable() => null;

        public Type GetFieldType(int i) => this[i].GetType();
        public string GetDataTypeName(int i) => this[i].GetType().Name;

        public bool IsDBNull(int i) => _recordsEnumerator!.Current[i] is null;
        public DateTime GetDateTime(int i)
        {
            if (this[i].GetType() == typeof(DateTime))
            {
                return (DateTime)this[i];
            }
            else
            {
                return DateTime.FromOADate(GetDouble(i));
            }
        }
        
        // Float values are stored in excel as double type
        // It's unlikely that the number is out of bounds, but if this happens, return zero.
        public decimal GetDecimal(int i)
        {
            try
            {
                return Convert.ToDecimal(this[i]);
            }
            catch (Exception)
            {
                return 0;
            }
        }
        
        // Remaining methods are just a conversion
        public double GetDouble(int i) => Convert.ToDouble(this[i]);
        public float GetFloat(int i) => (float)Convert.ToDouble(this[i]);
        public short GetInt16(int i) => Convert.ToInt16(this[i]);
        public int GetInt32(int i) => Convert.ToInt32(this[i]);
        public long GetInt64(int i) => Convert.ToInt64(this[i]);
        public string GetString(int i) => this?[i].ToString() ?? string.Empty;
        public object GetValue(int i) => this[i];


        public void Dispose()
        {
            
        }

        #endregion



        #region Private Methods

        private Dictionary<string, int> GetHeaderMap()
        {
            Dictionary<string, int> map = new();
            object?[] header = _records.First();

            int indent = 0;
            if (_records.GetType() == typeof(ExcelIterator))
            {
                ExcelIterator records = (ExcelIterator)_records;
                indent = (int)records.MinCol;
            }

            if (HeaderSource == HeaderSource.FirstRow)
            {
                for (int i = 0; i < header.Length; i++)
                {
                    string headerVal = header[i]?.ToString()
                        ?? "NoName" + i.ToString("D" + header.Length.ToString().Length);

                    map.Add(headerVal, i);
                }
                _records = _records.Skip(1);
            }

            if (HeaderSource == HeaderSource.Noheader)
            {
                for (int i = 0; i < header.Length; i++)
                {
                    map.Add(ExcelExtensions.IndexToLetter(i + indent), i);
                }
            }


            if (HeaderSource == HeaderSource.JsonMap)
            {
                if (JsonMap is null)
                    throw new NullReferenceException($"Parameter JsonMap is not set!");

                map = GetMapFromJson(JsonMap, header.Length);
            }

            if (HeaderSource == HeaderSource.DictMap)
            {
                if (DictMap is null)
                    throw new NullReferenceException($"Parameter DictMap is not set!");

                map = DictMap;
            }

            return map;
        }


        
        #endregion



        #region Not Supported Methods
        
        
        public IDataReader GetData(int i) => throw new NotSupportedException();
        public bool GetBoolean(int i) => throw new NotSupportedException();
        public byte GetByte(int i) => throw new NotSupportedException();
        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length) => throw new NotSupportedException();
        public char GetChar(int i) => throw new NotSupportedException();
        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length) => throw new NotSupportedException();
        public Guid GetGuid(int i) => throw new NotSupportedException();

        #endregion


    }
}
