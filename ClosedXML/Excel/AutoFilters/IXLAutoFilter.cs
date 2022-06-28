// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLFilterDynamicType { AboveAverage, BelowAverage }

    public enum XLFilterType { Regular, Custom, TopBottom, Dynamic, DateTimeGrouping }

    public enum XLTopBottomPart { Top, Bottom }

    public interface IXLAutoFilter
    {
        [Obsolete("Use IsEnabled")]
        Boolean Enabled { get; set; }
        IEnumerable<IXLRangeRow> HiddenRows { get; }
        Boolean IsEnabled { get; set; }
        IXLRange Range { get; set; }
        short SortColumn { get; set; }
        Boolean Sorted { get; set; }
        XLSortOrder SortOrder { get; set; }
        IEnumerable<IXLRangeRow> VisibleRows { get; }

        IXLAutoFilter Clear();

        IXLFilterColumn Column(String column);

        IXLFilterColumn Column(short column);

        IXLAutoFilter Reapply();

        IXLAutoFilter Sort(short columnToSortBy = 1, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);
    }
}
