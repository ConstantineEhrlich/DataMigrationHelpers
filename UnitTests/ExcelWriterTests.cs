using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelHelpers;
using System.IO;

namespace UnitTests
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class ExcelWriterTests
    {
        private readonly string exportPath = @"/Users/constantine/Documents/code/dotnet";

        private static IDataReader GetSampleData()
        {
            DataTable table = new();
            table.Columns.Add("Project", typeof(string));
            table.Columns.Add("Order Currency", typeof(string));
            table.Columns.Add("Currency Rate", typeof(decimal));
            table.Columns.Add("Order Amount", typeof(decimal));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add("A101", "USD",  1.0m,  258424.25m, new DateTime(2022, 8, 14));
            table.Rows.Add("A102", "ILS", 3.48m,  748515.78m, new DateTime(2021, 5, 3));
            table.Rows.Add("A103", "EUR", 0.95m, 4541831.15m, new DateTime(2018, 7, 7));
            table.Rows.Add("X572", "GBP", 0.83m,  3889651.2m, new DateTime(2017, 1, 15));
            table.Rows.Add("W445", "USD",  1.0m, 88183005.0m, new DateTime(2014, 9, 9));

            return table.CreateDataReader();
        }


        [TestMethod]
        public void ReadSampleData()
        {
            IDataReader reader = GetSampleData();
            while (reader.Read())
            {
                
            }
        }

        [TestMethod]
        public void WriteToExcel()
        {
            using ExcelWriter writer = new(Path.Combine(exportPath, "NewTest.xlsx"));

            writer.WriteData("Sheet Test", GetSampleData());
        }



    }
}
