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
    }
}
