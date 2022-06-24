#nullable enable
namespace ClosedXML.Excel
{
    internal class XLStyleKeyBuilder
    {
        private readonly XLStyleKey _styleKey;
        private XLProtectionKey? _protection;
        private XLAlignmentKey? _alignment;
        private XLBorderKey? _border;
        private XLFillKey? _fill;
        private XLFontKey? _font;
        private bool? _includeQuotePrefix;
        private XLNumberFormatKey? _numberFormat;

        public XLStyleKeyBuilder(XLStyleKey styleKey)
        {
            _styleKey = styleKey;
        }

        public XLStyleKeyBuilder WithProtection(XLProtectionKey protection)
        {
            _protection = protection;
            return this;
        }

        public XLStyleKeyBuilder WithAlignment(XLAlignmentKey alignment)
        {
            _alignment = alignment;
            return this;
        }

        public XLStyleKeyBuilder WithBorder(XLBorderKey border)
        {
            _border = border;
            return this;
        }

        public XLStyleKeyBuilder WithFill(XLFillKey fill)
        {
            _fill = fill;
            return this;
        }

        public XLStyleKeyBuilder WithFont(XLFontKey font)
        {
            _font = font;
            return this;
        }

        public XLStyleKeyBuilder WithIncludeQuotePrefix(bool includeQuotePrefix)
        {
            _includeQuotePrefix = includeQuotePrefix;
            return this;
        }

        public XLStyleKeyBuilder WithNumberFormat(XLNumberFormatKey numberFormat)
        {
            _numberFormat = numberFormat;
            return this;
        }

        public XLStyleKey Build()
        {
            return new XLStyleKey(_alignment ?? _styleKey.Alignment, _border ?? _styleKey.Border, _fill ?? _styleKey.Fill,
                _font ?? _styleKey.Font,
                _includeQuotePrefix ?? _styleKey.IncludeQuotePrefix,
                _numberFormat ?? _styleKey.NumberFormat, _protection ?? _styleKey.Protection);
        }
    }
}
