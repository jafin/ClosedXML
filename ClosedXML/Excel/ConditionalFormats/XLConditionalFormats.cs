using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLConditionalFormats : IXLConditionalFormats
    {
        private readonly List<IXLConditionalFormat> _conditionalFormats = new List<IXLConditionalFormat>();

        private static readonly List<XLConditionalFormatType> _conditionalFormatTypesExcludedFromConsolidation = new List<XLConditionalFormatType>()
        {
            XLConditionalFormatType.DataBar,
            XLConditionalFormatType.ColorScale,
            XLConditionalFormatType.IconSet,
            XLConditionalFormatType.Top10,
            XLConditionalFormatType.AboveAverage,
            XLConditionalFormatType.IsDuplicate,
            XLConditionalFormatType.IsUnique
        };

        public void Add(IXLConditionalFormat conditionalFormat)
        {
            _conditionalFormats.Add(conditionalFormat);
        }

        public IEnumerator<IXLConditionalFormat> GetEnumerator()
        {
            return _conditionalFormats.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Remove(Predicate<IXLConditionalFormat> predicate)
        {
            _conditionalFormats.RemoveAll(predicate);
        }


        // Get the top left corner of the rectangle covering all the ranges
        private (int row, int col) GetTopLeftFromRange(IXLRanges ranges)
        {
            var minRow = int.MaxValue;
            var minCol = int.MaxValue;
            foreach (var range in ranges)
            {
                if (range.RangeAddress.FirstAddress.RowNumber < minRow)
                    minRow = range.RangeAddress.FirstAddress.RowNumber;
                if (range.RangeAddress.FirstAddress.ColumnNumber < minCol)
                    minCol = range.RangeAddress.FirstAddress.ColumnNumber;
            }
            return (minRow, minCol);
        }

        /// <summary>
        /// The method consolidate the same conditional formats, which are located in adjacent ranges.
        /// </summary>
        internal void Consolidate()
        {
            var formats = _conditionalFormats
                .Where(cf => cf.Ranges.Any())
                .ToList();
            _conditionalFormats.Clear();

            while (formats.Count > 0)
            {
                var item = formats.First();

                if (!_conditionalFormatTypesExcludedFromConsolidation.Contains(item.ConditionalFormatType))
                {
                    var rangesToJoin = new XLRanges();
                    item.Ranges.ForEach(r => rangesToJoin.Add(r));
                    var firstRange = item.Ranges.First();
                    var skippedRanges = new XLRanges();
                    bool IsSameFormat(IXLConditionalFormat f) => f != item
                                                                 && f.Ranges.First().Worksheet.Position == firstRange.Worksheet.Position
                                                                 && XLConditionalFormat.NoRangeComparer.Equals(f, item);
                    var topLeftAddress = GetTopLeftFromRange(item.Ranges);
                    var baseAddress = new XLAddress(topLeftAddress.row,topLeftAddress.col, false, false);
                    var baseCell = firstRange.Worksheet.Cell(baseAddress) as XLCell;

                    int i = 1;
                    bool stop = false;
                    List<IXLConditionalFormat> similarFormats = new List<IXLConditionalFormat>();
                    do
                    {
                        stop = (i >= formats.Count);

                        if (!stop)
                        {
                            var nextFormat = formats[i];

                            var intersectsSkipped =
                                skippedRanges.Any(left => nextFormat.Ranges.GetIntersectedRanges(left.RangeAddress).Any());

                            var isSameFormat = IsSameFormat(nextFormat);

                            if (isSameFormat && !intersectsSkipped)
                            {
                                similarFormats.Add(nextFormat);
                                nextFormat.Ranges.ForEach(r => rangesToJoin.Add(r));
                            }
                            else if (rangesToJoin.Any(left => nextFormat.Ranges.GetIntersectedRanges(left.RangeAddress).Any()) ||
                                     intersectsSkipped)
                            {
                                // if we reached the rule intersecting any of captured ranges stop for not breaking the priorities
                                stop = true;
                            }

                            if (!isSameFormat)
                                nextFormat.Ranges.ForEach(r => skippedRanges.Add(r));
                        }

                        i++;
                    } while (!stop);

                    var consRanges = rangesToJoin.Consolidate();
                    item.Ranges.RemoveAll();
                    consRanges.ForEach(r => item.Ranges.Add(r));

                    var targetCell = item.Ranges.First().FirstCell() as XLCell;
                    (item as XLConditionalFormat).AdjustFormulas(baseCell, targetCell);

                    similarFormats.ForEach(cf => formats.Remove(cf));
                }

                _conditionalFormats.Add(item);
                formats.Remove(item);
            }
        }

        public void RemoveAll()
        {
            _conditionalFormats.Clear();
        }

        /// <summary>
        /// Reorders the according to original priority. Done during load process
        /// </summary>
        public void ReorderAccordingToOriginalPriority()
        {
            var reorderedFormats = _conditionalFormats.OrderBy(cf => (cf as XLConditionalFormat).OriginalPriority).ToList();
            _conditionalFormats.Clear();
            _conditionalFormats.AddRange(reorderedFormats);
        }
    }
}
