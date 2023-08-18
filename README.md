# DataMigrationHelpers

A collection of tools and utilities to assist with data migration tasks,
focusing on Excel operations and data serialization.

## Contents

### Excel helpers
Multi-plaform tools that allow excel files reading and writing.
Utilizes the ```DocumentFormat.OpenXml``` library
- **ExcelIterator.cs** - A class that reads the contents of the Excel file.
- **ExcelExtensions.cs** - Extension methods to convert excel literals into numbers, and vise-versa.
- **DataReaderFactory.cs** - A factory class that creates ```IDataReader``` from different data structures
  (for example, fom ExcelIterator or a collection of dictionaries).
- **ExcelWriter.cs** - A class that writes the contents of the ```IDataReader``` into the excel file and saves it on the disk.
- **CreateExcelStylesheet.cs** - Helper class that adds a stylesheet into the excel file.

### Serialization
Utilizes ```System.Reflection```
- **Serializer.cs** - A class that allows serialization of objects into array of objects, or into the ```Dictionary<string, object>```.
- **Deserializer.cs** - Provides generic method ```Deserialize<T>``` that creates an objects from ```IDataReader``` class.

### Unit tests
Unit tests for Excel files reading and writing, and for serialization and deserialization.


## Contribution
Feel free to contribute to this project by submitting pull requests or reporting issues.

## License
This project is licensed under the MIT License.