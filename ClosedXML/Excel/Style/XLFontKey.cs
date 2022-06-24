#nullable enable
using System;

namespace ClosedXML.Excel
{
    internal readonly struct XLFontKey : IEquatable<XLFontKey>
    {
        public XLFontKey(bool bold, bool italic, XLFontUnderlineValues underline, bool strikethrough,
            XLFontVerticalTextAlignmentValues verticalAlignment, bool shadow, double fontSize, XLColorKey fontColor, string fontName,
            XLFontFamilyNumberingValues fontFamilyNumbering, XLFontCharSet fontCharSet) : this()
        {
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strikethrough = strikethrough;
            VerticalAlignment = verticalAlignment;
            Shadow = shadow;
            FontSize = fontSize;
            FontColor = fontColor;
            FontName = fontName;
            FontFamilyNumbering = fontFamilyNumbering;
            FontCharSet = fontCharSet;
            _hashCode = CalculateHashCode();
        }

        public bool Bold { get; }

        public bool Italic { get; }

        public XLFontUnderlineValues Underline { get; }

        public bool Strikethrough { get; }

        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; }

        public bool Shadow { get; }

        public double FontSize { get; }

        public XLColorKey FontColor { get; }

        public string FontName { get; }

        public XLFontFamilyNumberingValues FontFamilyNumbering { get; }

        public XLFontCharSet FontCharSet { get; }
        private readonly int _hashCode;

        public bool Equals(XLFontKey other)
        {
            return
                Bold == other.Bold
                && Italic == other.Italic
                && Underline == other.Underline
                && Strikethrough == other.Strikethrough
                && VerticalAlignment == other.VerticalAlignment
                && Shadow == other.Shadow
                && FontSize == other.FontSize
                && FontColor == other.FontColor
                && FontFamilyNumbering == other.FontFamilyNumbering
                && FontCharSet == other.FontCharSet
                && string.Equals(FontName, other.FontName, StringComparison.InvariantCultureIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            if (obj is XLFontKey key)
                return Equals(key);
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        private int CalculateHashCode()
        {
            var hashCode = 1158783753;
            hashCode = hashCode * -1521134295 + Bold.GetHashCode();
            hashCode = hashCode * -1521134295 + Italic.GetHashCode();
            hashCode = hashCode * -1521134295 + (int)Underline;
            hashCode = hashCode * -1521134295 + Strikethrough.GetHashCode();
            hashCode = hashCode * -1521134295 + (int)VerticalAlignment;
            hashCode = hashCode * -1521134295 + Shadow.GetHashCode();
            hashCode = hashCode * -1521134295 + FontSize.GetHashCode();
            hashCode = hashCode * -1521134295 + FontColor.GetHashCode();
            hashCode = hashCode * -1521134295 + StringComparer.InvariantCultureIgnoreCase.GetHashCode(FontName);
            hashCode = hashCode * -1521134295 + (int)FontFamilyNumbering;
            hashCode = hashCode * -1521134295 + (int)FontCharSet;
            return hashCode;
        }

        public override string ToString()
        {
            return $"{FontName} {FontSize}pt {FontColor} " +
                   (Bold ? "Bold" : "") + (Italic ? "Italic" : "") + (Strikethrough ? "Strikethrough" : "") +
                   (Underline == XLFontUnderlineValues.None ? "" : Underline.ToString()) +
                   $"{FontFamilyNumbering} {FontCharSet}";
        }

        public static bool operator ==(XLFontKey left, XLFontKey right) => left.Equals(right);

        public static bool operator !=(XLFontKey left, XLFontKey right) => !(left.Equals(right));
    }
}
