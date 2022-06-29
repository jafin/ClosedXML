namespace ClosedXML.Excel
{
    public struct ColumnVal
    {
        public ColumnVal(short value)
        {
            this.Value = value;
        }

        public ColumnVal(int value)
        {
            this.Value = (short)value;
        }

        public short Value { get; set; }

        public bool Equals(ColumnVal other)
        {
            return Value == other.Value;
        }

        public override bool Equals(object obj)
        {
            return obj is ColumnVal other && Equals(other);
        }

        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }
    }
}
