using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLAddressComparer : IEqualityComparer<IXLAddress>
    {
        private readonly bool _ignoreFixed;

        public XLAddressComparer(bool ignoreFixed)
        {
            _ignoreFixed = ignoreFixed;
        }

        public bool Equals(IXLAddress x, IXLAddress y)
        {
            return (x == null && y == null) ||
                   (x != null && y != null &&
                    string.Equals(x.Worksheet.Name, y.Worksheet.Name, StringComparison.InvariantCultureIgnoreCase) &&
                    x.ColumnNumber == y.ColumnNumber &&
                    x.RowNumber == y.RowNumber &&
                    (_ignoreFixed || x.FixedColumn == y.FixedColumn &&
                     x.FixedRow == y.FixedRow));
        }

        public int GetHashCode(IXLAddress obj)
        {
            var hashCode = new HashCode();
            hashCode.Add(obj.Worksheet.Name, StringComparer.InvariantCultureIgnoreCase);
            hashCode.Add(obj.ColumnNumber);
            hashCode.Add(obj.RowNumber);
            hashCode.Add(!_ignoreFixed && obj.FixedColumn);
            hashCode.Add(!_ignoreFixed && obj.FixedRow);
            return hashCode.ToHashCode();
        }
    }
}
