using System.Runtime.InteropServices;

namespace EmailParser.Helpers;

/// <summary>
/// Utility methods for safely releasing COM objects.
/// </summary>
internal static class ComHelper
{
    /// <summary>
    /// Releases a COM object reference if it is not <c>null</c> and is a valid COM object.
    /// </summary>
    internal static void ReleaseComObject(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
            Marshal.ReleaseComObject(comObject);
    }
}
