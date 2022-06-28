using System;

namespace ClosedXML.Excel
{
    internal readonly struct XLSheetPoint:IEquatable<XLSheetPoint>
    {
        public XLSheetPoint(Int32  row, short column)
        {
            Row = row;
            Column = column;
        }

        public readonly Int32 Row;
        public readonly short Column;

        public override bool Equals(object obj)
        {
            return Equals((XLSheetPoint)obj);
        }

        public bool Equals(XLSheetPoint other)
        {
            return Row == other.Row && Column == other.Column;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Row, Column);
        }

        public static bool operator==(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row == b.Row && a.Column == b.Column;
        }

        public static bool operator !=(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row != b.Row || a.Column != b.Column;
        }
    }
}
