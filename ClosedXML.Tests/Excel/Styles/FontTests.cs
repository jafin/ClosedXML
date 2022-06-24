using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class FontTests
    {
        private XLFontKey GetDefaultFontKey()
        {
            return new XLFontKey(false, false, XLFontUnderlineValues.None, false,
                XLFontVerticalTextAlignmentValues.Baseline, false, 11,
                XLColor.FromArgb(0, 0, 0).Key, "Calibri", XLFontFamilyNumberingValues.Swiss, XLFontCharSet.Default);
        }

        [Test]
        public void XLFontKey_GetHashCode_IsCaseInsensitive()
        {
            var fontKey1 = new XLFontKeyBuilder(GetDefaultFontKey()).WithFontName("Arial").Build();
            var fontKey2 = new XLFontKeyBuilder(GetDefaultFontKey()).WithFontName("Times New Roman").Build();
            var fontKey3 = new XLFontKeyBuilder(GetDefaultFontKey()).WithFontName("TIMES NEW ROMAN").Build();

            Assert.AreNotEqual(fontKey1.GetHashCode(), fontKey2.GetHashCode());
            Assert.AreEqual(fontKey2.GetHashCode(), fontKey3.GetHashCode());
        }

        [Test]
        public void XLFontKey_Equals_IsCaseInsensitive()
        {
            var fontKey1 = new XLFontKeyBuilder(GetDefaultFontKey()).WithFontName("Arial").Build();
            var fontKey2 = new XLFontKeyBuilder(GetDefaultFontKey()).WithFontName("Times New Roman").Build();
            var fontKey3 = new XLFontKeyBuilder(GetDefaultFontKey()).WithFontName("TIMES NEW ROMAN").Build();

            Assert.IsFalse(fontKey1.Equals(fontKey2));
            Assert.IsTrue(fontKey2.Equals(fontKey3));
        }
    }
}
