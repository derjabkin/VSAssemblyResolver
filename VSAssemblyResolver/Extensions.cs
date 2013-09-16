using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SergejDerjabkin.VSAssemblyResolver
{
    static class Extensions
    {
        internal static IEnumerable<T> Where<T>(this IList<T> list, Predicate<T> predicate, int startIndex, int endIndex)
        {
            if (list == null) throw new ArgumentNullException("list");
            if (predicate == null) throw new ArgumentNullException("predicate");

            for (int i = startIndex; i <= endIndex; i++)
            {
                if (predicate(list[i])) yield return list[i];
            }
        }

        internal static IEnumerable<T> ToEnumerable<T>(this T item)
        {
            yield return item;
        }
    }
}
