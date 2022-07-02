using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    public readonly struct XLRangeAddress : IEquatable<XLRangeAddress>
    {
        #region Static members

        public static XLRangeAddress EntireColumn(IXLWorksheet worksheet, int column)
        {
            return EntireColumn(worksheet, (short)column);
        }

        public static XLRangeAddress EntireColumn(IXLWorksheet worksheet, short column)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, 1, column, false, false),
                new XLAddress(worksheet, XLHelper.MaxRowNumber, column, false, false));
        }

        public static XLRangeAddress EntireRow(IXLWorksheet worksheet, int row)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, row, 1, false, false),
                new XLAddress(worksheet, row, XLHelper.MaxColumnNumber, false, false));
        }

        public static readonly XLRangeAddress Invalid = new XLRangeAddress(
            new XLAddress(-1, -1, fixedRow: true, fixedColumn: true),
            new XLAddress(-1, -1, fixedRow: true, fixedColumn: true)
        );

        #endregion Static members

        #region Constructor

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress) : this()
        {
            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(IXLWorksheet worksheet, string rangeAddress) : this()
        {
            string addressToUse = rangeAddress.Contains("!")
                ? rangeAddress.Substring(rangeAddress.LastIndexOf("!", StringComparison.Ordinal) + 1)
                : rangeAddress;

            string firstPart;
            string secondPart;
            if (addressToUse.Contains(':'))
            {
                var arrRange = addressToUse.Split(':');
                firstPart = arrRange[0];
                secondPart = arrRange[1];
            }
            else
            {
                firstPart = addressToUse;
                secondPart = addressToUse;
            }

            if (XLHelper.IsValidA1Address(firstPart))
            {
                FirstAddress = XLAddress.Create(worksheet, firstPart);
                LastAddress = XLAddress.Create(worksheet, secondPart);
            }
            else
            {
                firstPart = firstPart.Replace("$", string.Empty);
                secondPart = secondPart.Replace("$", string.Empty);
                if (char.IsDigit(firstPart[0]))
                {
                    FirstAddress = XLAddress.Create(worksheet, "A" + firstPart);
                    LastAddress = XLAddress.Create(worksheet, XLHelper.MaxColumnLetter + secondPart);
                }
                else
                {
                    FirstAddress = XLAddress.Create(worksheet, firstPart + "1");
                    LastAddress = XLAddress.Create(worksheet, secondPart + XLHelper.MaxRowNumber.ToInvariantString());
                }
            }

            Worksheet = worksheet;
        }

        #endregion Constructor

        #region Public properties

        public IXLWorksheet Worksheet { get; }

        /// <summary>
        /// Gets or sets the first address in the range.
        /// </summary>
        /// <value>
        /// The first address.
        /// </value>
        public XLAddress FirstAddress { get; }

        /// <summary>
        /// Gets or sets the last address in the range.
        /// </summary>
        /// <value>
        /// The last address.
        /// </value>
        public XLAddress LastAddress { get; }

        /// <summary>
        /// Gets or sets a value indicating whether this range is valid.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is valid; otherwise, <c>false</c>.
        /// </value>
        public bool IsValid => FirstAddress.IsValid && LastAddress.IsValid;

        /// <summary>
        /// Gets the number of columns in the area covered by the range address.
        /// </summary>
        public int ColumnSpan
        {
            get
            {
                if (!IsValid)
                    throw new InvalidOperationException("Range address is invalid.");

                return Math.Abs(LastAddress.ColumnNumber - FirstAddress.ColumnNumber) + 1;
            }
        }

        /// <summary>
        /// Gets the number of cells in the area covered by the range address.
        /// </summary>
        public int NumberOfCells => ColumnSpan * RowSpan;

        /// <summary>
        /// Gets the number of rows in the area covered by the range address.
        /// </summary>
        public int RowSpan
        {
            get
            {
                if (!IsValid)
                    throw new InvalidOperationException("Range address is invalid.");

                return Math.Abs(LastAddress.RowNumber - FirstAddress.RowNumber) + 1;
            }
        }

        private bool WorksheetIsDeleted => Worksheet?.IsDeleted == true;

        #endregion Public properties

        #region Public methods

        public Boolean IsNormalized => LastAddress.RowNumber >= FirstAddress.RowNumber
                                       && LastAddress.ColumnNumber >= FirstAddress.ColumnNumber;

        /// <summary>
        /// Lead a range address to a normal form - when <see cref="FirstAddress"/> points to the top-left address and
        /// <see cref="LastAddress"/> points to the bottom-right address.
        /// </summary>
        /// <returns></returns>
        public XLRangeAddress Normalize()
        {
            if (FirstAddress.RowNumber <= LastAddress.RowNumber &&
                FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
                return this;

            int firstRow, lastRow;
            short firstColumn, lastColumn;
            bool firstRowFixed, firstColumnFixed, lastRowFixed, lastColumnFixed;

            if (FirstAddress.RowNumber <= LastAddress.RowNumber)
            {
                firstRow = FirstAddress.RowNumber;
                firstRowFixed = FirstAddress.FixedRow;
                lastRow = LastAddress.RowNumber;
                lastRowFixed = LastAddress.FixedRow;
            }
            else
            {
                firstRow = LastAddress.RowNumber;
                firstRowFixed = LastAddress.FixedRow;
                lastRow = FirstAddress.RowNumber;
                lastRowFixed = FirstAddress.FixedRow;
            }

            if (FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
            {
                firstColumn = FirstAddress.ColumnNumber;
                firstColumnFixed = FirstAddress.FixedColumn;
                lastColumn = LastAddress.ColumnNumber;
                lastColumnFixed = LastAddress.FixedColumn;
            }
            else
            {
                firstColumn = LastAddress.ColumnNumber;
                firstColumnFixed = LastAddress.FixedColumn;
                lastColumn = FirstAddress.ColumnNumber;
                lastColumnFixed = FirstAddress.FixedColumn;
            }

            return new XLRangeAddress(
                new XLAddress(FirstAddress.Worksheet, firstRow, firstColumn, firstRowFixed, firstColumnFixed),
                new XLAddress(LastAddress.Worksheet, lastRow, lastColumn, lastRowFixed, lastColumnFixed));
        }

        /// <summary>
        /// See if the two ranges intersect
        /// </summary>
        /// <param name="otherAddress"></param>
        /// <returns></returns>
        public bool Intersects(XLRangeAddress otherAddress)
        {
            return !(
                       otherAddress.FirstAddress.ColumnNumber > LastAddress.ColumnNumber
                    || otherAddress.LastAddress.ColumnNumber < FirstAddress.ColumnNumber
                    || otherAddress.FirstAddress.RowNumber > LastAddress.RowNumber
                    || otherAddress.LastAddress.RowNumber < FirstAddress.RowNumber
                );
        }

        internal XLRangeAddress WithoutWorksheet()
        {
            return new XLRangeAddress(
                FirstAddress.WithoutWorksheet(),
                LastAddress.WithoutWorksheet());
        }

        public bool Contains(XLAddress address)
        {
            return FirstAddress.RowNumber <= address.RowNumber &&
                   address.RowNumber <= LastAddress.RowNumber &&
                   FirstAddress.ColumnNumber <= address.ColumnNumber &&
                   address.ColumnNumber <= LastAddress.ColumnNumber;
        }

        public String ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public String ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.A1);
        }

        public String ToStringRelative(Boolean includeSheet)
        {
            string address;
            if (!IsValid)
                address = "#REF!";
            else
            {
                if (IsEntireSheet())
                    address = $"1:{XLHelper.MaxRowNumber}";
                else if (IsEntireRow())
                    address = String.Concat(FirstAddress.RowNumber.ToString(), ":", LastAddress.RowNumber.ToString());
                else if (IsEntireColumn())
                    address = String.Concat(FirstAddress.ColumnLetter, ":", LastAddress.ColumnLetter);
                else
                    address = String.Concat(FirstAddress.ToStringRelative(), ":", LastAddress.ToStringRelative());
            }

            if (includeSheet || WorksheetIsDeleted)
                return String.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    "!", address);

            return address;
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
        {
            string address;
            if (!IsValid)
                address = "#REF!";
            else
            {
                if (IsEntireSheet())
                    address = $"$1:${XLHelper.MaxRowNumber}";
                else if (IsEntireRow())
                    address = String.Concat("$", FirstAddress.RowNumber.ToString(), ":$", LastAddress.RowNumber.ToString());
                else if (IsEntireColumn())
                    address = String.Concat("$", FirstAddress.ColumnLetter, ":$", LastAddress.ColumnLetter);
                else
                    address = String.Concat(FirstAddress.ToStringFixed(referenceStyle), ":", LastAddress.ToStringFixed(referenceStyle));
            }

            if (includeSheet || WorksheetIsDeleted)
                return String.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    "!", address);

            return address;
        }

        public override string ToString()
        {
            if (!IsValid || WorksheetIsDeleted)
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";

                var address = (!FirstAddress.IsValid || !LastAddress.IsValid) ? "#REF!" : String.Concat(FirstAddress.ToString(), ":", LastAddress.ToString());
                return String.Concat(worksheet, address);
            }

            if (IsEntireSheet())
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";
                var address = $"1:{XLHelper.MaxRowNumber}";
                return String.Concat(worksheet, address);
            }
            else if (IsEntireRow())
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";
                var firstAddress = FirstAddress.IsValid ? FirstAddress.RowNumber.ToString() : "#REF!";
                var lastAddress = LastAddress.IsValid ? LastAddress.RowNumber.ToString() : "#REF!";

                return String.Concat(worksheet, firstAddress, ':', lastAddress);
            }
            else if (IsEntireColumn())
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";
                var firstAddress = FirstAddress.IsValid ? FirstAddress.ColumnLetter : "#REF!";
                var lastAddress = LastAddress.IsValid ? LastAddress.ColumnLetter : "#REF!";

                return String.Concat(worksheet, firstAddress, ':', lastAddress);
            }
            else
            {
                return String.Concat(FirstAddress.ToString(), ":", LastAddress.ToString());
            }
        }

        public string ToString(XLReferenceStyle referenceStyle)
        {
            return ToString(referenceStyle, false);
        }

        public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            if (referenceStyle == XLReferenceStyle.R1C1)
                return ToStringFixed(referenceStyle, true);
            else
                return ToStringRelative(includeSheet);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is XLRangeAddress address))
            {
                return false;
            }

            return FirstAddress.Equals(address.FirstAddress) &&
                   LastAddress.Equals(address.LastAddress) &&
                   EqualityComparer<IXLWorksheet>.Default.Equals(Worksheet, address.Worksheet);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FirstAddress, LastAddress) * -1521134295 + EqualityComparer<IXLWorksheet>.Default.GetHashCode(Worksheet);
        }

        public bool Equals(XLRangeAddress other)
        {
            return ReferenceEquals(Worksheet, other.Worksheet) &&
                   FirstAddress == other.FirstAddress &&
                   LastAddress == other.LastAddress;
        }

        /// <summary>
        /// Determines whether range address spans the entire column.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire column; otherwise, <c>false</c>.
        /// </returns>        
        public bool IsEntireColumn()
        {
            return IsValid
                   && FirstAddress.RowNumber == 1
                   && LastAddress.RowNumber == XLHelper.MaxRowNumber;
        }

        /// <summary>
        /// Determines whether range address spans the entire row.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire row; otherwise, <c>false</c>.
        /// </returns>
        public bool IsEntireRow()
        {
            return IsValid
                   && FirstAddress.ColumnNumber == 1
                   && LastAddress.ColumnNumber == XLHelper.MaxColumnNumber;
        }

        /// <summary>
        /// Determines whether the range address spans the entire worksheet.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if is entire sheet; otherwise, <c>false</c>.
        /// </returns>        
        public bool IsEntireSheet()
        {
            return IsValid && IsEntireColumn() && IsEntireRow();
        }

        /// <summary>
        /// Returns a range address so that its offset from the target base address is equal to the offset of the current range address to the source base address.
        /// For example, if the current range address is D4:E4, the source base address is A1:C3, then the relative address to the target base address B10:D13 is E14:F14
        /// </summary>
        /// <param name="sourceRangeAddress">The source base range address.</param>
        /// <param name="targetRangeAddress">The target base range address.</param>
        /// <returns>The relative range</returns>        
        public XLRangeAddress Relative(XLRangeAddress sourceRangeAddress, XLRangeAddress targetRangeAddress)
        {
            var sheet = targetRangeAddress.Worksheet;

            return new XLRangeAddress
            (
                new XLAddress
                (
                    sheet,
                    this.FirstAddress.RowNumber - sourceRangeAddress.FirstAddress.RowNumber + targetRangeAddress.FirstAddress.RowNumber,
                    (short)(this.FirstAddress.ColumnNumber - sourceRangeAddress.FirstAddress.ColumnNumber + targetRangeAddress.FirstAddress.ColumnNumber),
                    fixedRow: false,
                    fixedColumn: false
                ),
                new XLAddress
                (
                    sheet,
                    this.LastAddress.RowNumber - sourceRangeAddress.FirstAddress.RowNumber + targetRangeAddress.FirstAddress.RowNumber,
                    (short)(this.LastAddress.ColumnNumber - sourceRangeAddress.FirstAddress.ColumnNumber + targetRangeAddress.FirstAddress.ColumnNumber),
                    fixedRow: false,
                    fixedColumn: false
                )
            );
        }

        /// <summary>
        /// Returns the intersection of this range address with another range address on the same worksheet.
        /// </summary>
        /// <param name="otherRangeAddress">The other range address.</param>
        /// <returns>The intersection's range address</returns>        
        public XLRangeAddress Intersection(XLRangeAddress otherRangeAddress)
        {
            if (!this.Worksheet.Equals(otherRangeAddress.Worksheet))
                throw new ArgumentOutOfRangeException(nameof(otherRangeAddress), "The other range address is on a different worksheet");

            var thisRangeAddressNormalized = this.Normalize();
            var otherRangeAddressNormalized = otherRangeAddress.Normalize();

            var firstRow = Math.Max(thisRangeAddressNormalized.FirstAddress.RowNumber, otherRangeAddressNormalized.FirstAddress.RowNumber);
            var firstColumn = Math.Max(thisRangeAddressNormalized.FirstAddress.ColumnNumber, otherRangeAddressNormalized.FirstAddress.ColumnNumber);
            var lastRow = Math.Min(thisRangeAddressNormalized.LastAddress.RowNumber, otherRangeAddressNormalized.LastAddress.RowNumber);
            var lastColumn = Math.Min(thisRangeAddressNormalized.LastAddress.ColumnNumber, otherRangeAddressNormalized.LastAddress.ColumnNumber);

            if (lastRow < firstRow || lastColumn < firstColumn)
                return XLRangeAddress.Invalid;

            return new XLRangeAddress
            (
                new XLAddress(this.Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
                new XLAddress(this.Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
            );
        }

        /// <summary>Allocates the current range address in the internal range repository and returns it</summary>
        public IXLRange AsRange()
        {
            if (this.Worksheet == null)
                throw new InvalidOperationException("The worksheet of the current range address has not been set.");

            if (!this.IsValid)
                return null;

            return this.Worksheet.Range(this);
        }

        #endregion Public methods

        #region Operators

        public static bool operator ==(XLRangeAddress left, XLRangeAddress right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(XLRangeAddress left, XLRangeAddress right)
        {
            return !(left == right);
        }

        #endregion Operators
    }
}
