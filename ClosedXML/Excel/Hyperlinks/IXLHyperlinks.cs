using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLHyperlinks: IEnumerable<XLHyperlink>
    {
        void Add(XLHyperlink hyperlink);
        void Delete(XLHyperlink hyperlink);
        void Delete(XLAddress address);
        bool TryDelete(XLAddress address);
        XLHyperlink Get(XLAddress address);
        bool TryGet(XLAddress address, out XLHyperlink hyperlink);
    }
}
