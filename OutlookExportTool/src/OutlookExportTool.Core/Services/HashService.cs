using System.Security.Cryptography;

namespace OutlookExportTool.Core.Services;

public sealed class HashService
{
    public string ComputeSha256(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}
