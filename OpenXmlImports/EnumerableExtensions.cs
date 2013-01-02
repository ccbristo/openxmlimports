using System.Collections.Generic;
using System.Linq;

namespace OpenXmlImports
{
    public static class EnumerableExtensions
    {
        public static bool In<T>(this T item, params T[] options)
        {
            return In(item, (IEnumerable<T>)options);
        }

        public static bool In<T>(this T item, IEnumerable<T> options)
        {
            return options.Any(option => (option == null && item == null) ||
                (option != null && option.Equals(item)));
        }
    }
}
