using System.Text.Json.Serialization;

namespace OutlookExportTool.Core.Models;

public sealed class ManifestEntry
{
    [JsonPropertyName("entry_id")] public string EntryId { get; set; } = string.Empty;
    [JsonPropertyName("export_status")] public string ExportStatus { get; set; } = "complete";
    [JsonPropertyName("canonical")] public CanonicalRecord Canonical { get; set; } = new();
    [JsonPropertyName("headers_path")] public string? HeadersPath { get; set; }
    [JsonPropertyName("attachments")] public List<AttachmentRecord> Attachments { get; set; } = new();
    [JsonPropertyName("source_metadata")] public EmailSourceMetadata SourceMetadata { get; set; } = new();
    [JsonPropertyName("error")] public string? Error { get; set; }
}

public sealed class CanonicalRecord
{
    [JsonPropertyName("format")] public string Format { get; set; } = "msg";
    [JsonPropertyName("path")] public string Path { get; set; } = string.Empty;
    [JsonPropertyName("hash_sha256")] public string HashSha256 { get; set; } = string.Empty;
    [JsonPropertyName("size_bytes")] public long SizeBytes { get; set; }
}

public sealed class AttachmentRecord
{
    [JsonPropertyName("path")] public string Path { get; set; } = string.Empty;
    [JsonPropertyName("original_filename")] public string OriginalFilename { get; set; } = string.Empty;
    [JsonPropertyName("hash_sha256")] public string HashSha256 { get; set; } = string.Empty;
    [JsonPropertyName("size_bytes")] public long SizeBytes { get; set; }
}

public sealed class EmailSourceMetadata
{
    [JsonPropertyName("folder_path")] public string FolderPath { get; set; } = string.Empty;
    [JsonPropertyName("subject")] public string? Subject { get; set; }
    [JsonPropertyName("sender")] public string? Sender { get; set; }
    [JsonPropertyName("recipients")] public List<string> Recipients { get; set; } = new();
    [JsonPropertyName("sent_utc")] public string? SentUtc { get; set; }
    [JsonPropertyName("received_utc")] public string? ReceivedUtc { get; set; }
    [JsonPropertyName("message_id")] public string? MessageId { get; set; }
}
