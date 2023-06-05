using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelHelpers;
using System.Collections;
using System.IO;



namespace UnitTests
{
    [TestClass]
    public class ExcelIteratorTests
    {
        public readonly string xlFilePath = @"/Users/constantine/Documents/code/dotnet/Employees.xlsx";
        public readonly string jsonFilePath = @"/Users/constantine/Documents/code/dotnet/EmployeeExcelMap.json";


        [TestMethod]
        public void ExcelIterator()
        {
            ExcelIterator et = new(xlFilePath);

            et.WorksheetName = "Employees";
            et.MinRow = 2;
            et.MaxRow = 25;
            et.MinCol = (uint)ExcelExtensions.LetterToIndex("B");
            et.MaxCol = (uint)ExcelExtensions.LetterToIndex("G");

            int i = 0;

            foreach (var item in et)
            {
                i++;
                Console.WriteLine(item[2]);
            }
        }



        [TestMethod]
        public void RowIndexTest()
        {
            ExcelIterator et = new(xlFilePath)
            {
                WorksheetName = "Employees"
            };

            Console.WriteLine(et.RowIndex);
        }

        [TestMethod]
        public void MinColLessThanMaxCol()
        {
            ExcelIterator et = new(xlFilePath);
            et.MaxCol = 10;
            Assert.ThrowsException<IndexOutOfRangeException>(() => et.MinCol = 15);
        }

        [TestMethod]
        public void MaxColGreaterThanMinCol()
        {
            ExcelIterator et = new(xlFilePath);
            et.MinCol = 10;
            Assert.ThrowsException<IndexOutOfRangeException>(() => et.MaxCol = 4);
        }




        [TestMethod]
        public void ExcelReaderTest()
        {
            string jsonMap = File.ReadAllText(jsonFilePath);

            ExcelIterator xlData = new(xlFilePath)
            {
                MinRow = 2,
                MaxRow = 22,
                MinCol = 1,
                MaxCol = 7,
                WorksheetName = "Employees"
            };


            DataReaderFactory reader = new(xlData, HeaderSource.JsonMap, jsonMap);

            while (reader.Read())
            {

            }




        }

    }
}