using System;

namespace ClosedXML.Report.Utils;

internal static class EmptyArray
{
    public static T[] New<T>()
    {
#if NETSTANDARD1_3_OR_GREATER
        return Array.Empty<T>();
#else
        return new T[0];
#endif
    }
}
