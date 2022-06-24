#nullable enable
namespace ClosedXML.Excel
{
    internal class XLFontKeyBuilder
    {
        private XLFontKey _fontKey;
        private bool? _bold;
        private bool? _italic;
        private XLFontUnderlineValues? _underline;
        private bool? _strikethrough;
        private XLFontVerticalTextAlignmentValues? _verticalAlignment;
        private bool? _shadow;
        private double? fontSize;
        private XLColorKey? _fontColor;
        private string? _fontName;
        private XLFontFamilyNumberingValues? _fontFamilyNumbering;
        private XLFontCharSet? _fontCharSet;

        public XLFontKeyBuilder(XLFontKey fontKey)
        {
            _fontKey = fontKey;
        }

        public XLFontKeyBuilder WithFontCharSet(XLFontCharSet fontCharSet)
        {
            _fontCharSet = fontCharSet;
            return this;
        }

        public XLFontKeyBuilder WithFontFamilyNumbering(XLFontFamilyNumberingValues fontFamilyNumbering)
        {
            _fontFamilyNumbering = fontFamilyNumbering;
            return this;
        }


        public XLFontKeyBuilder WithFontName(string fontName)
        {
            _fontName = fontName;
            return this;
        }

        public XLFontKeyBuilder WithFontColor(XLColorKey fontColor)
        {
            _fontColor = fontColor;
            return this;
        }

        public XLFontKeyBuilder WithFontSize(double fontSize)
        {
            this.fontSize = fontSize;
            return this;
        }

        public XLFontKeyBuilder WithShadow(bool shadow)
        {
            _shadow = shadow;
            return this;
        }

        public XLFontKeyBuilder WithVerticalAlignment(XLFontVerticalTextAlignmentValues verticalAlignment)
        {
            _verticalAlignment = verticalAlignment;
            return this;
        }

        public XLFontKeyBuilder WithStrikethrough(bool strikeThrough)
        {
            _strikethrough = strikeThrough;
            return this;
        }


        public XLFontKeyBuilder WithUnderline(XLFontUnderlineValues underline)
        {
            _underline = underline;
            return this;
        }

        public XLFontKeyBuilder WithItalic(bool italic)
        {
            _italic = italic;
            return this;
        }

        public XLFontKeyBuilder WithBold(bool bold)
        {
            _bold = bold;
            return this;
        }

        public XLFontKey Build()
        {
            return new XLFontKey(_bold ?? _fontKey.Bold, _italic ?? _fontKey.Italic, _underline ?? _fontKey.Underline,
                _strikethrough ?? _fontKey.Strikethrough, _verticalAlignment ?? _fontKey.VerticalAlignment, _shadow ?? _fontKey.Shadow,
                fontSize ?? _fontKey.FontSize, _fontColor ?? _fontKey.FontColor, _fontName ?? _fontKey.FontName,
                _fontFamilyNumbering ?? _fontKey.FontFamilyNumbering, _fontCharSet ?? _fontKey.FontCharSet);
        }
    }
}
