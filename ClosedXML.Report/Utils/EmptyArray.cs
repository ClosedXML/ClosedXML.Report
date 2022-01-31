using System;

namespace ClosedXML.Report.Utils;

internal static class EmptyArray
{
    public static T[] New<T>()
    {
#if NETSTANDARD
        return Array.Empty<T>();
#else
        return new T[0];
#endif
    }
}
