using OutlookExportTool.Core.Utilities;
using Xunit;

namespace OutlookExportTool.Tests;

public class SanitizerTests
{
    [Fact]
    public void SanitizesInvalidCharacters()
    {
        var name = FileNameSanitizer.Sanitize("report<>.txt");
        Assert.DoesNotContain('<', name);
        Assert.DoesNotContain('>', name);
    }
}
