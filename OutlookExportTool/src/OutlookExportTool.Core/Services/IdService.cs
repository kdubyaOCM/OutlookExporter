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
        var sanitizedSender = sender ?? string.Empty;
        var sanitizedSubject = subject ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(messageId))
        {
            return JoinFields(messageId.Trim(), sanitizedSender, sanitizedSubject, time, body);
        }

        if (searchKeyFallback is { Length: > 0 })
        {
            var fallbackHex = Convert.ToHexString(searchKeyFallback);
            return JoinFields(fallbackHex, sanitizedSender, sanitizedSubject, time, body);
        }

        return JoinFields(entryIdFallback ?? string.Empty, sanitizedSender, sanitizedSubject, time, body);
    }

    private static string JoinFields(params string[] fields)
    {
        return string.Join("|", fields);
    }
}
