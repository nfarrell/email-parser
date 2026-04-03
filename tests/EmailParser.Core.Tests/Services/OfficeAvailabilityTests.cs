using System.Runtime.InteropServices;
using EmailParser.Core.Services;

namespace EmailParser.Core.Tests.Services;

public class OfficeAvailabilityTests
{
    // ───────────────────────── Direct exception types ─────────────────────────

    [Fact]
    public void IsOfficeUnavailableException_TypeLoadException_ReturnsTrue()
    {
        var ex = new TypeLoadException("Could not load type");
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_FileNotFoundException_ReturnsTrue()
    {
        var ex = new FileNotFoundException("Assembly not found");
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_FileLoadException_ReturnsTrue()
    {
        var ex = new FileLoadException("Could not load assembly");
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_BadImageFormatException_ReturnsTrue()
    {
        var ex = new BadImageFormatException("Bad image format");
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_COMExceptionRegDbClassNotReg_ReturnsTrue()
    {
        // REGDB_E_CLASSNOTREG = 0x80040154
        var ex = new COMException("Class not registered", unchecked((int)0x80040154u));
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    // ───────────────────────── Negative cases ─────────────────────────

    [Fact]
    public void IsOfficeUnavailableException_GenericException_ReturnsFalse()
    {
        var ex = new Exception("Something went wrong");
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_InvalidOperationException_ReturnsFalse()
    {
        var ex = new InvalidOperationException("Invalid operation");
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_ArgumentException_ReturnsFalse()
    {
        var ex = new ArgumentException("Bad argument");
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_COMExceptionOtherHResult_ReturnsFalse()
    {
        // Some other HRESULT that is not REGDB_E_CLASSNOTREG
        var ex = new COMException("Other COM error", unchecked((int)0x80004005u));
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    [Fact]
    public void IsOfficeUnavailableException_NullInnerException_ReturnsFalse()
    {
        var ex = new IOException("IO error");
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(ex));
    }

    // ───────────────────────── Inner exception recursion ─────────────────────────

    [Fact]
    public void IsOfficeUnavailableException_TypeLoadExceptionAsInner_ReturnsTrue()
    {
        var inner = new TypeLoadException("Could not load type");
        var outer = new InvalidOperationException("Wrapper", inner);
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(outer));
    }

    [Fact]
    public void IsOfficeUnavailableException_FileNotFoundExceptionAsDeepInner_ReturnsTrue()
    {
        var innermost = new FileNotFoundException("Assembly not found");
        var middle = new InvalidOperationException("Middle", innermost);
        var outer = new Exception("Outer wrapper", middle);
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(outer));
    }

    [Fact]
    public void IsOfficeUnavailableException_COMExceptionAsInner_ReturnsTrue()
    {
        var inner = new COMException("Class not registered", unchecked((int)0x80040154u));
        var outer = new Exception("Wrapper", inner);
        Assert.True(OfficeAvailability.IsOfficeUnavailableException(outer));
    }

    [Fact]
    public void IsOfficeUnavailableException_NoMatchingInnerException_ReturnsFalse()
    {
        var inner = new ArgumentException("Bad arg");
        var outer = new InvalidOperationException("Wrapper", inner);
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(outer));
    }

    [Fact]
    public void IsOfficeUnavailableException_DeeplyNestedNoMatch_ReturnsFalse()
    {
        var innermost = new IOException("IO error");
        var middle = new ArgumentException("Bad arg", innermost);
        var outer = new Exception("Outer", middle);
        Assert.False(OfficeAvailability.IsOfficeUnavailableException(outer));
    }
}
