using System;

namespace ClosedXML.Excel
{
    public struct XLAlignmentKey : IEquatable<XLAlignmentKey>
    {
        public XLAlignmentHorizontalValues Horizontal { get; set; }

        public XLAlignmentVerticalValues Vertical { get; set; }

        public int Indent { get; set; }

        public bool JustifyLastLine { get; set; }

        public XLAlignmentReadingOrderValues ReadingOrder { get; set; }

        public int RelativeIndent { get; set; }

        public bool ShrinkToFit { get; set; }

        public int TextRotation { get; set; }

        public bool WrapText { get; set; }

        public bool TopToBottom { get; set; }

        public bool Equals(XLAlignmentKey other)
        {
            return
                Horizontal == other.Horizontal
             && Vertical == other.Vertical
             && Indent == other.Indent
             && JustifyLastLine == other.JustifyLastLine
             && ReadingOrder == other.ReadingOrder
             && RelativeIndent == other.RelativeIndent
             && ShrinkToFit == other.ShrinkToFit
             && TextRotation == other.TextRotation
             && WrapText == other.WrapText
             && TopToBottom == other.TopToBottom;
        }

        public override bool Equals(object obj)
        {
            if (obj is XLAlignment)
                return Equals((XLAlignment)obj);
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            var hashCode = new HashCode();
            hashCode.Add(Horizontal);
            hashCode.Add(Vertical);
            hashCode.Add(Indent);
            hashCode.Add(JustifyLastLine);
            hashCode.Add(ReadingOrder);
            hashCode.Add(RelativeIndent);
            hashCode.Add(ShrinkToFit);
            hashCode.Add(TextRotation);
            hashCode.Add(WrapText);
            hashCode.Add(TopToBottom);
            return hashCode.ToHashCode();
        }

        public override string ToString()
        {
            return
                $"{Horizontal} {Vertical} {ReadingOrder} Indent: {Indent} RelativeIndent: {RelativeIndent} TextRotation: {TextRotation} " +
                (WrapText ? "WrapText" : "") +
                (JustifyLastLine ? "JustifyLastLine" : "") +
                (TopToBottom ? "TopToBottom" : "");
        }

        public static bool operator ==(XLAlignmentKey left, XLAlignmentKey right) => left.Equals(right);

        public static bool operator !=(XLAlignmentKey left, XLAlignmentKey right) => !(left.Equals(right));
    }
}
