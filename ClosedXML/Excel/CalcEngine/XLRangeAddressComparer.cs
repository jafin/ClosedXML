using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLRangeAddressComparer : IEqualityComparer<XLRangeAddress>
    {
        private readonly XLAddressComparer _addressComparer;

        public XLRangeAddressComparer(bool ignoreFixed)
        {
            _addressComparer = new XLAddressComparer(ignoreFixed);
        }

        public bool Equals(XLRangeAddress x, XLRangeAddress y)
        {
            return (x == null && y == null) ||
                   (x != null && y != null &&
                    _addressComparer.Equals(x.FirstAddress, y.FirstAddress) &&
                    _addressComparer.Equals(x.LastAddress, y.LastAddress));
        }

        public int GetHashCode(XLRangeAddress obj)
        {
            return HashCode.Combine(_addressComparer.GetHashCode(obj.FirstAddress), _addressComparer.GetHashCode(obj.LastAddress));
        }
    }
}
