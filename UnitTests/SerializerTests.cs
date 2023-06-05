using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Serialization;

namespace UnitTests
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class SerializerTests
    {
        [TestMethod]
        public void TestGetPropertyMap()
        {
            Dictionary<string, int> map = Serializer.GetPropertyMap(typeof(DateTime));
        }
    }
}
