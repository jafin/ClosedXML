using System;
using NUnit.Framework;

namespace ClosedXML.Tests
{
    public class SpanTest
    {
        [Test]
        public void X()
        {
            string z = "HelloWorld";
            ReadOnlySpan<char> x = z.AsSpan();
            var zz = x.Slice(2).ToString();
            int i = 5;
        }

        [TestCase("'B1!B2")]
        [TestCase("B1!B2")]
        public void TestTarget(string target)
        {
            var result = GetTargetCell(target);
            var result2 = GetTargetCell(target);
            Assert.AreEqual(result,result2);
        }
        
        public string GetTargetCell(string target)
        {
            var pair = target.Split('!');
            if (pair.Length == 1)
                return "1";

            var wsName = pair[0].AsSpan();
            if (wsName.StartsWith("'".AsSpan()))
                wsName = wsName.Slice(1, wsName.Length - 2);
            return wsName.ToString();
        }

        public string GetTargetCellOriginal(string target)
        {
            var pair = target.Split('!');
            if (pair.Length == 1)
                return "1";

            var wsName = pair[0];
            if (wsName.StartsWith("'"))
                wsName = wsName.Substring(1, wsName.Length - 2);
            return wsName.ToString();
        }
    }
}
