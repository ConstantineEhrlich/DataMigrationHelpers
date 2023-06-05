using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTests
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class ExcelReaderFromDict
    {
        public List<Dictionary<string, object?>> SampleData()
        {
            return new List<Dictionary<string, object?>>()
            {
                new Dictionary<string, object?>()
                {
                    { "Project", "X382" },
                    { "Date Created", new DateTime(2022, 12, 1) },
                    { "Money", 5884.548m }
                },

                new Dictionary<string, object?>()
                {
                    { "Project", "X386" },
                    { "Date Created", new DateTime(2018, 5, 13) },
                    { "Money", 2884.57m }
                },

                new Dictionary<string, object?>()
                {
                    { "Project", "X884" },
                    { "Date Created", new DateTime(2014, 3, 8) },
                    { "Money", 9985.57m }
                },
            };

        }



        [TestMethod]
        public void DictMapTest()
        {
            Dictionary<string, object?> first = SampleData().First();

            Dictionary<string, int> map = first.Select((kv, index) => new { k = kv.Key, v = index }).ToDictionary(kv => kv.k, kv => kv.v);

            IEnumerable<object?[]> vals = SampleData().Select(dict => dict.Values.ToArray());
        }


        [TestMethod]
        public void MakeXlRd()
        {
            ExcelHelpers.DataReaderFactory reader = new(SampleData());
            while (reader.Read())
            {


            }
        }


        [TestMethod]
        public void MakeXlWrt()
        {
            string filePath = @"/Users/constantine/Documents/code/dotnet/test.xlsx";

            using ExcelHelpers.ExcelWriter writer = new(filePath);

            ExcelHelpers.DataReaderFactory reader = new(SampleData());

            writer.WriteData("Test output", reader);

        }
    }
}
