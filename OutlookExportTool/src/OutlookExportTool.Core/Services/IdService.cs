using System.Security.Cryptography;
using System.Text;
using OutlookExportTool.Core.Utilities;

namespace OutlookExportTool.Core.Services;

public sealed class IdService
{
    public string CreateDeterministicId(string? messageId, string? sender, string? subject, DateTime? sentOrReceivedUtc, string? bodyPreview, byte[]? searchKeyFallback, string? entryIdFallback)
    {
        var input = BuildInput(messageId, sender, subject, sentOrReceivedUtc, bodyPreview, searchKeyFallback, entryIdFallback);
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(input));
        Span<byte> first16 = stackalloc byte[16];
        hash.AsSpan(0, 16).CopyTo(first16);
        var base32 = Base32.Encode(first16);
        return base32;
    }

    private static string BuildInput(string? messageId, string? sender, string? subject, DateTime? sentOrReceivedUtc, string? bodyPreview, byte[]? searchKeyFallback, string? entryIdFallback)
    {
        var body = bodyPreview ?? string.Empty;
        var time = sentOrReceivedUtc?.ToString("o") ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(messageId))
        {
            return string.Join("|", messageId.Trim(), sender ?? string.Empty, subject ?? string.Empty, time, body);
        }

        if (searchKeyFallback is { Length: > 0 })
        {
            var fallbackHex = Convert.ToHexString(searchKeyFallback);
            return string.Join("|", fallbackHex, sender ?? string.Empty, subject ?? string.Empty, time, body);
        }

        return string.Join("|", entryIdFallback ?? string.Empty, sender ?? string.Empty, subject ?? string.Empty, time, body);
    }
}
