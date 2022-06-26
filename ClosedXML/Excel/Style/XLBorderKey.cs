using System;

namespace ClosedXML.Excel
{
    internal struct XLBorderKey : IEquatable<XLBorderKey>
    {
        public XLBorderStyleValues LeftBorder { get; set; }

        public XLColorKey LeftBorderColor { get; set; }

        public XLBorderStyleValues RightBorder { get; set; }

        public XLColorKey RightBorderColor { get; set; }

        public XLBorderStyleValues TopBorder { get; set; }

        public XLColorKey TopBorderColor { get; set; }

        public XLBorderStyleValues BottomBorder { get; set; }

        public XLColorKey BottomBorderColor { get; set; }

        public XLBorderStyleValues DiagonalBorder { get; set; }

        public XLColorKey DiagonalBorderColor { get; set; }

        public bool DiagonalUp { get; set; }

        public bool DiagonalDown { get; set; }

        public override int GetHashCode()
        {
            var hashCode = new HashCode();
            hashCode.Add(LeftBorder);
            hashCode.Add(RightBorder);
            hashCode.Add(TopBorder);
            hashCode.Add(BottomBorder);
            hashCode.Add(DiagonalBorder);
            hashCode.Add(DiagonalUp);
            hashCode.Add(DiagonalDown);

            if (LeftBorder != XLBorderStyleValues.None)
                hashCode.Add(LeftBorderColor);
            if (RightBorder != XLBorderStyleValues.None)
                hashCode.Add(RightBorderColor);
            if (TopBorder != XLBorderStyleValues.None)
                hashCode.Add(TopBorderColor);
            if (BottomBorder != XLBorderStyleValues.None)
                hashCode.Add(BottomBorderColor);
            if (DiagonalBorder != XLBorderStyleValues.None)
                hashCode.Add(DiagonalBorderColor);
            return hashCode.ToHashCode();
        }

        public bool Equals(XLBorderKey other)
        {
            return
                   AreEquivalent(LeftBorder, LeftBorderColor, other.LeftBorder, other.LeftBorderColor)
                && AreEquivalent(RightBorder, RightBorderColor, other.RightBorder, other.RightBorderColor)
                && AreEquivalent(TopBorder, TopBorderColor, other.TopBorder, other.TopBorderColor)
                && AreEquivalent(BottomBorder, BottomBorderColor, other.BottomBorder, other.BottomBorderColor)
                && AreEquivalent(DiagonalBorder, DiagonalBorderColor, other.DiagonalBorder, other.DiagonalBorderColor)
                && DiagonalUp == other.DiagonalUp
                && DiagonalDown == other.DiagonalDown;
        }

        private bool AreEquivalent(
            XLBorderStyleValues borderStyle1, XLColorKey color1,
            XLBorderStyleValues borderStyle2, XLColorKey color2)
        {
            return (borderStyle1 == XLBorderStyleValues.None &&
                    borderStyle2 == XLBorderStyleValues.None) ||
                   borderStyle1 == borderStyle2 &&
                   color1 == color2;
        }

        public override bool Equals(object obj)
        {
            if (obj is XLBorderKey)
                return Equals((XLBorderKey)obj);
            return base.Equals(obj);
        }

        public override string ToString()
        {
            return $"{LeftBorder} {LeftBorderColor} {RightBorder} {RightBorderColor} {TopBorder} {TopBorderColor} " +
                   $"{BottomBorder} {BottomBorderColor} {DiagonalBorder} {DiagonalBorderColor} " +
                   (DiagonalUp ? "DiagonalUp" : "") +
                   (DiagonalDown ? "DiagonalDown" : "");
        }

        public static bool operator ==(XLBorderKey left, XLBorderKey right) => left.Equals(right);

        public static bool operator !=(XLBorderKey left, XLBorderKey right) => !(left.Equals(right));
    }
}
