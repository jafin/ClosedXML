using System;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLStyle : IXLStyle
    {
        #region Static members

        public static XLStyle Default
        {
            get { return new XLStyle(XLStyleValue.Default); }
        }

        internal static XLStyleKey GenerateKey(IXLStyle initialStyle)
        {
            if (initialStyle == null)
                return Default.Key;
            if (initialStyle is XLStyle style)
                return style.Key;

            return new XLStyleKey(XLAlignment.GenerateKey(initialStyle.Alignment),
                XLBorder.GenerateKey(initialStyle.Border),
                XLFill.GenerateKey(initialStyle.Fill),
                XLFont.GenerateKey(initialStyle.Font), default(bool), XLNumberFormat.GenerateKey(initialStyle.NumberFormat),
                XLProtection.GenerateKey(initialStyle.Protection)
            );
        }

        internal static XLStyle CreateEmptyStyle()
        {
            return new XLStyle(new XLStylizedEmpty(null));
        }

        #endregion Static members

        #region properties

        private readonly IXLStylized _container;

        internal XLStyleValue Value { get; private set; }

        internal XLStyleKey Key
        {
            get { return Value.Key; }
            private set { Value = XLStyleValue.FromKey(ref value); }
        }

        #endregion properties

        #region constructors

        public XLStyle(IXLStylized container, IXLStyle initialStyle = null, Boolean useDefaultModify = true) : this(container,
            GenerateKey(initialStyle))
        {
        }

        public XLStyle(IXLStylized container, XLStyleKey key) : this(container, XLStyleValue.FromKey(ref key))
        {
        }

        internal XLStyle(IXLStylized container, XLStyleValue value)
        {
            _container = container ?? new XLStylizedEmpty(XLStyle.Default);
            Value = value;
        }

        /// <summary>
        /// To initialize XLStyle.Default only
        /// </summary>
        private XLStyle(XLStyleValue value)
        {
            _container = null;
            Value = value;
        }

        #endregion constructors

        internal void Modify(Func<XLStyleKey, XLStyleKey> modification)
        {
            Key = modification(Key);

            if (_container != null)
            {
                _container.ModifyStyle(modification);
            }
        }

        #region IXLStyle members

        public IXLFont Font
        {
            get { return new XLFont(this, Value.Font); }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithFont(XLFont.GenerateKey(value)).Build());
            }
        }

        public IXLAlignment Alignment
        {
            get { return new XLAlignment(this, Value.Alignment); }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithAlignment(XLAlignment.GenerateKey(value)).Build());
            }
        }

        public IXLBorder Border
        {
            get { return new XLBorder(_container, this, Value.Border); }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithBorder(XLBorder.GenerateKey(value)).Build());
            }
        }

        public IXLFill Fill
        {
            get { return new XLFill(this, Value.Fill); }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithFill(XLFill.GenerateKey(value)).Build());
            }
        }

        public Boolean IncludeQuotePrefix
        {
            get { return Value.IncludeQuotePrefix; }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithIncludeQuotePrefix(value).Build());
            }
        }

        public IXLStyle SetIncludeQuotePrefix(Boolean includeQuotePrefix = true)
        {
            IncludeQuotePrefix = includeQuotePrefix;
            return this;
        }

        public IXLNumberFormat NumberFormat
        {
            get { return new XLNumberFormat(this, Value.NumberFormat); }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithNumberFormat(XLNumberFormat.GenerateKey(value)).Build());
            }
        }

        public IXLProtection Protection
        {
            get { return new XLProtection(this, Value.Protection); }
            set
            {
                Modify(k => new XLStyleKeyBuilder(k).WithProtection(XLProtection.GenerateKey(value)).Build());
            }
        }

        public IXLNumberFormat DateFormat
        {
            get { return NumberFormat; }
        }

        #endregion IXLStyle members

        #region Overridden

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("Font:");
            sb.Append(Font);
            sb.Append(" Fill:");
            sb.Append(Fill);
            sb.Append(" Border:");
            sb.Append(Border);
            sb.Append(" NumberFormat: ");
            sb.Append(NumberFormat);
            sb.Append(" Alignment: ");
            sb.Append(Alignment);
            sb.Append(" Protection: ");
            sb.Append(Protection);
            return sb.ToString();
        }

        public bool Equals(IXLStyle other)
        {
            var otherS = other as XLStyle;

            if (otherS == null)
                return false;

            return Key == otherS.Key &&
                   _container == otherS._container;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XLStyle);
        }

        public override int GetHashCode()
        {
            var hashCode = 416600561;
            hashCode = hashCode * -1521134295 + Key.GetHashCode();
            return hashCode;
        }

        #endregion Overridden
    }
}
