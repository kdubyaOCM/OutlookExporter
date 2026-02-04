using OutlookExportTool.Core.Models;
using OutlookExportTool.Core.Services;
using Xunit;

namespace OutlookExportTool.Tests;

public class ManifestServiceTests
{
    [Fact]
    public void AppendsAndLoadsManifest()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "manifest-tests");
        Directory.CreateDirectory(tempDir);
        var manifestPath = Path.Combine(tempDir, "manifest.jsonl");
        if (File.Exists(manifestPath))
        {
            File.Delete(manifestPath);
        }

        var service = new ManifestService();
        var entry = new ManifestEntry
        {
            EntryId = "ABC",
            ExportStatus = "complete",
            Canonical = new CanonicalRecord { Path = "emails/ABC/ABC.msg", HashSha256 = "hash", SizeBytes = 12 }
        };

        service.Append(manifestPath, entry);
        var index = service.LoadIndex(manifestPath);

        Assert.True(index.ContainsKey("ABC"));
    }
}
