using System.Runtime.InteropServices;

namespace EmailParser.Services;

/// <summary>
/// Helpers for detecting whether a Microsoft Office component is available
/// on the current machine.
/// </summary>
internal static class OfficeAvailability
{
    /// <summary>
    /// Returns <see langword="true"/> when <paramref name="ex"/> indicates that a
    /// required Office assembly or COM component could not be loaded — i.e. Office is
    /// not installed or not properly registered.
    /// </summary>
    /// <remarks>
    /// Covers the following failure modes:
    /// <list type="bullet">
    ///   <item><see cref="TypeLoadException"/> — JIT failed to load an Office interop type.</item>
    ///   <item><see cref="FileNotFoundException"/> / <see cref="FileLoadException"/> — an Office
    ///     assembly (e.g. <c>office.dll</c>) could not be found.</item>
    ///   <item><see cref="BadImageFormatException"/> — assembly architecture mismatch.</item>
    ///   <item><see cref="COMException"/> HRESULT <c>0x80040154</c> (REGDB_E_CLASSNOTREG) —
    ///     the COM class is not registered because Office is not installed.</item>
    /// </list>
    /// The check recurses into <see cref="Exception.InnerException"/> so that wrapper
    /// exceptions (e.g. <see cref="InvalidOperationException"/>) are also matched.
    /// </remarks>
    internal static bool IsOfficeUnavailableException(Exception ex)
    {
        if (ex is TypeLoadException or FileNotFoundException or FileLoadException
                or BadImageFormatException)
            return true;

        // REGDB_E_CLASSNOTREG (0x80040154): COM class is not registered — Office not installed.
        if (ex is COMException comEx && (uint)comEx.HResult == 0x80040154u)
            return true;

        return ex.InnerException is not null && IsOfficeUnavailableException(ex.InnerException);
    }
}
