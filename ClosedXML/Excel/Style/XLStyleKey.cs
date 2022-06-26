#nullable enable
using System;

namespace ClosedXML.Excel
{
    internal readonly struct XLStyleKey : IEquatable<XLStyleKey>
    {
        public XLStyleKey(XLAlignmentKey alignment, XLBorderKey border, XLFillKey fill, XLFontKey font, bool includeQuotePrefix,
            XLNumberFormatKey numberFormat, XLProtectionKey protection) : this()
        {
            Alignment = alignment;
            Border = border;
            Fill = fill;
            Font = font;
            IncludeQuotePrefix = includeQuotePrefix;
            NumberFormat = numberFormat;
            Protection = protection;
            _hashCode = CalculateHashCode();
        }

        public XLAlignmentKey Alignment { get; }

        public XLBorderKey Border { get; }

        public XLFillKey Fill { get; }

        public XLFontKey Font { get; }

        public bool IncludeQuotePrefix { get; }

        public XLNumberFormatKey NumberFormat { get; }

        public XLProtectionKey Protection { get; }

        private readonly int _hashCode;

        public override int GetHashCode()
        {
            return _hashCode;
        }

        private int CalculateHashCode()
        {
            return HashCode.Combine(Alignment, Border, Fill, Font, IncludeQuotePrefix, NumberFormat, Protection);
        }

        public bool Equals(XLStyleKey other)
        {
            return Alignment == other.Alignment &&
                   Border == other.Border &&
                   Fill == other.Fill &&
                   Font == other.Font &&
                   IncludeQuotePrefix == other.IncludeQuotePrefix &&
                   NumberFormat == other.NumberFormat &&
                   Protection == other.Protection;
        }

        public override string ToString()
        {
            return
                this == XLStyle.Default.Key
                    ? "Default"
                    : $"Alignment: {(Alignment == XLStyle.Default.Key.Alignment ? "Default" : Alignment.ToString())} Border: {(Border == XLStyle.Default.Key.Border ? "Default" : Border.ToString())} Fill: {(Fill == XLStyle.Default.Key.Fill ? "Default" : Fill.ToString())} Font: {(Font == XLStyle.Default.Key.Font ? "Default" : Font.ToString())} IncludeQuotePrefix: {(IncludeQuotePrefix == XLStyle.Default.Key.IncludeQuotePrefix ? "Default" : IncludeQuotePrefix.ToString())} NumberFormat: {(NumberFormat == XLStyle.Default.Key.NumberFormat ? "Default" : NumberFormat.ToString())} Protection: {(Protection == XLStyle.Default.Key.Protection ? "Default" : Protection.ToString())}";
        }

        public override bool Equals(object obj)
        {
            if (obj is XLStyleKey key)
                return Equals(key);
            return base.Equals(obj);
        }

        public static bool operator ==(XLStyleKey left, XLStyleKey right) => left.Equals(right);

        public static bool operator !=(XLStyleKey left, XLStyleKey right) => !(left.Equals(right));

        public void Deconstruct(
            out XLAlignmentKey alignment,
            out XLBorderKey border,
            out XLFillKey fill,
            out XLFontKey font,
            out Boolean includeQuotePrefix,
            out XLNumberFormatKey numberFormat,
            out XLProtectionKey protection)
        {
            alignment = Alignment;
            border = Border;
            fill = Fill;
            font = Font;
            includeQuotePrefix = IncludeQuotePrefix;
            numberFormat = NumberFormat;
            protection = Protection;
        }
    }
}
