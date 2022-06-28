using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLColumnsCollection : IDictionary<short, XLColumn>
    {
        public void ShiftColumnsRight(short startingColumn, short columnsToShift)
        {
            foreach (var co in _dictionary.Keys.Where(k => k >= startingColumn).OrderByDescending(k => k))
            {
                var columnToMove = _dictionary[co];
                _dictionary.Remove(co);
                short newColumnNum = (short)(co + columnsToShift);
                if (newColumnNum <= XLHelper.MaxColumnNumber)
                {
                    columnToMove.SetColumnNumber(newColumnNum);
                    _dictionary.Add(newColumnNum, columnToMove);
                }
            }
        }

        private readonly Dictionary<short, XLColumn> _dictionary = new Dictionary<short, XLColumn>();

        public void Add(short key, XLColumn value)
        {
            _dictionary.Add(key, value);
        }

        public bool ContainsKey(short key)
        {
            return _dictionary.ContainsKey(key);
        }

        public ICollection<short> Keys
        {
            get { return _dictionary.Keys; }
        }

        public bool Remove(short key)
        {
            return _dictionary.Remove(key);
        }

        public bool TryGetValue(short key, out XLColumn value)
        {
            return _dictionary.TryGetValue(key, out value);
        }

        public ICollection<XLColumn> Values
        {
            get { return _dictionary.Values; }
        }

        public XLColumn this[short key]
        {
            get
            {
                return _dictionary[key];
            }
            set
            {
                _dictionary[key] = value;
            }
        }

        public void Add(KeyValuePair<short, XLColumn> item)
        {
            _dictionary.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            _dictionary.Clear();
        }

        public bool Contains(KeyValuePair<short, XLColumn> item)
        {
            return _dictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<short, XLColumn>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return _dictionary.Count; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(KeyValuePair<short, XLColumn> item)
        {
            return _dictionary.Remove(item.Key);
        }

        public IEnumerator<KeyValuePair<short, XLColumn>> GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _dictionary.GetEnumerator();
        }

        public void RemoveAll(Func<XLColumn, Boolean> predicate)
        {
            _dictionary.RemoveAll(predicate);
        }
    }
}
