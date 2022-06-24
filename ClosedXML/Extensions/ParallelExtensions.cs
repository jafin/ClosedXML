using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClosedXML.Extensions
{
    public static class ParallelExtensions
    {
        public static IEnumerable<T1> OrderedParallel<T, T1>(this IEnumerable<T> list, Func<T, T1> action)
        {
            var unorderedResult = new ConcurrentBag<(long, T1)>();
            Parallel.ForEach(list, (o, state, i) => { unorderedResult.Add((i, action.Invoke(o))); });
            var ordered = unorderedResult.OrderBy(o => o.Item1);
            return ordered.Select(o => o.Item2);
        }
    }
}
