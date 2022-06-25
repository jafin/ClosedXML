using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLFont : IXLFont
    {
        #region Static members

        public static IXLFontBase DefaultCommentFont
        {
            get
            {
                // MS Excel uses Tahoma 9 Swiss no matter what current style font
                var defaultCommentFont = new XLFont
                {
                    FontName = "Tahoma",
                    FontSize = 9,
                    FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss
                };

                return defaultCommentFont;
            }
        }

        internal static XLFontKey GenerateKey(IXLFontBase defaultFont)
        {
            return defaultFont switch
            {
                null => XLFontValue.Default.Key,
                XLFont font => font.Key,
                _ => new XLFontKey(defaultFont.Bold, defaultFont.Italic, defaultFont.Underline, defaultFont.Strikethrough,
                    defaultFont.VerticalAlignment, defaultFont.Shadow, defaultFont.FontSize, defaultFont.FontColor.Key, defaultFont.FontName,
                    defaultFont.FontFamilyNumbering, defaultFont.FontCharSet)
            };
        }

        #endregion Static members

        private readonly XLStyle _style;

        private XLFontValue _value;

        internal XLFontKey Key
        {
            get { return _value.Key; }
            private set { _value = XLFontValue.FromKey(ref value); }
        }

        #region Constructors

        /// <summary>
        /// Create an instance of XLFont initializing it with the specified value.
        /// </summary>
        /// <param name="style">Style to attach the new instance to.</param>
        /// <param name="value">Style value to use.</param>
        public XLFont(XLStyle style, XLFontValue value)
        {
            _style = style ?? XLStyle.CreateEmptyStyle();
            _value = value;
        }

        public XLFont(XLStyle style, XLFontKey key) : this(style, XLFontValue.FromKey(ref key))
        {
        }

        public XLFont(XLStyle style = null, IXLFont d = null) : this(style, GenerateKey(d))
        {
        }

        #endregion Constructors

        private void Modify(Func<XLFontKey, XLFontKey> modification)
        {
            Key = modification(Key);

            _style.Modify(styleKey =>
            {
                var font = styleKey.Font;
                return new XLStyleKeyBuilder(styleKey).WithFont(modification(font)).Build();
            });
        }

        #region IXLFont Members

        public Boolean Bold
        {
            get { return Key.Bold; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithBold(value).Build());
            }
        }

        public Boolean Italic
        {
            get { return Key.Italic; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithItalic(value).Build());
            }
        }

        public XLFontUnderlineValues Underline
        {
            get { return Key.Underline; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithUnderline(value).Build());
            }
        }

        public Boolean Strikethrough
        {
            get { return Key.Strikethrough; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithStrikethrough(value).Build());
            }
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get { return Key.VerticalAlignment; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithVerticalAlignment(value).Build());
            }
        }

        public Boolean Shadow
        {
            get { return Key.Shadow; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithShadow(value).Build());
            }
        }

        public Double FontSize
        {
            get { return Key.FontSize; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithFontSize(value).Build());
            }
        }

        public XLColor FontColor
        {
            get
            {
                var fontColorKey = Key.FontColor;
                return XLColor.FromKey(ref fontColorKey);
            }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value), "Color cannot be null");
                Modify(k => new XLFontKeyBuilder(k).WithFontColor(value.Key).Build());
            }
        }

        public String FontName
        {
            get { return Key.FontName; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithFontName(value).Build());
            }
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get { return Key.FontFamilyNumbering; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithFontFamilyNumbering(value).Build());
            }
        }

        public XLFontCharSet FontCharSet
        {
            get { return Key.FontCharSet; }
            set
            {
                Modify(k => new XLFontKeyBuilder(k).WithFontCharSet(value).Build());
            }
        }

        public IXLStyle SetBold()
        {
            Bold = true;
            return _style;
        }

        public IXLStyle SetBold(Boolean value)
        {
            Bold = value;
            return _style;
        }

        public IXLStyle SetItalic()
        {
            Italic = true;
            return _style;
        }

        public IXLStyle SetItalic(Boolean value)
        {
            Italic = value;
            return _style;
        }

        public IXLStyle SetUnderline()
        {
            Underline = XLFontUnderlineValues.Single;
            return _style;
        }

        public IXLStyle SetUnderline(XLFontUnderlineValues value)
        {
            Underline = value;
            return _style;
        }

        public IXLStyle SetStrikethrough()
        {
            Strikethrough = true;
            return _style;
        }

        public IXLStyle SetStrikethrough(Boolean value)
        {
            Strikethrough = value;
            return _style;
        }

        public IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
        {
            VerticalAlignment = value;
            return _style;
        }

        public IXLStyle SetShadow()
        {
            Shadow = true;
            return _style;
        }

        public IXLStyle SetShadow(Boolean value)
        {
            Shadow = value;
            return _style;
        }

        public IXLStyle SetFontSize(Double value)
        {
            FontSize = value;
            return _style;
        }

        public IXLStyle SetFontColor(XLColor value)
        {
            FontColor = value;
            return _style;
        }

        public IXLStyle SetFontName(String value)
        {
            FontName = value;
            return _style;
        }

        public IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
        {
            FontFamilyNumbering = value;
            return _style;
        }

        public IXLStyle SetFontCharSet(XLFontCharSet value)
        {
            FontCharSet = value;
            return _style;
        }

        #endregion IXLFont Members

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Bold.ToString());
            sb.Append("-");
            sb.Append(Italic.ToString());
            sb.Append("-");
            sb.Append(Underline.ToString());
            sb.Append("-");
            sb.Append(Strikethrough.ToString());
            sb.Append("-");
            sb.Append(VerticalAlignment.ToString());
            sb.Append("-");
            sb.Append(Shadow.ToString());
            sb.Append("-");
            sb.Append(FontSize.ToString());
            sb.Append("-");
            sb.Append(FontColor);
            sb.Append("-");
            sb.Append(FontName);
            sb.Append("-");
            sb.Append(FontFamilyNumbering.ToString());
            return sb.ToString();
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XLFont);
        }

        public Boolean Equals(IXLFont other)
        {
            var otherF = other as XLFont;
            if (otherF == null)
                return false;

            return Key == otherF.Key;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = 416600561;
                return hashCode * -1521134295 + Key.GetHashCode();
            }
        }

        #endregion Overridden
    }
}
