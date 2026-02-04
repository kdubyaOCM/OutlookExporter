using OutlookExportTool.Core.Services;
using Xunit;

namespace OutlookExportTool.Tests;

public class IdServiceTests
{
    [Fact]
    public void DeterministicId_IsStable()
    {
        var service = new IdService();
        var id1 = service.CreateDeterministicId("<mid>", "sender", "subject", new DateTime(2024, 1, 1, 12, 0, 0, DateTimeKind.Utc), "body", null, null);
        var id2 = service.CreateDeterministicId("<mid>", "sender", "subject", new DateTime(2024, 1, 1, 12, 0, 0, DateTimeKind.Utc), "body", null, null);

        Assert.Equal(id1, id2);
        Assert.Equal(26, id1.Length);
    }
}
