using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using OutlookExportTool.Core.Models;
using OutlookExportTool.Core.Services;
using OutlookExportTool.Core.Utilities;
using OutlookExportTool.Outlook;

namespace OutlookExportTool.WPF;

public sealed class OutlookExportWorker
{
    private readonly HashService _hashService = new();
    private readonly IdService _idService = new();
    private readonly ManifestService _manifestService = new();
    private readonly StateService _stateService = new();

    public ExportSummary Run(ExportOptions options, Action<ProgressInfo> progress, CancellationToken token)
    {
        var summary = new ExportSummary();
        var exportRoot = ResolveExportRoot(options, out var resumeState);
        summary.ExportRoot = exportRoot;

        Directory.CreateDirectory(exportRoot);
        var logsPath = Path.Combine(exportRoot, "logs", $"export_{DateTime.Now:yyyyMMdd}.log");
        var log = new LogService(logsPath);

        var manifestPath = Path.Combine(exportRoot, "manifest.jsonl");
        var statePath = Path.Combine(exportRoot, "state.json");
        var emailsRoot = Path.Combine(exportRoot, "emails");
        Directory.CreateDirectory(emailsRoot);

        log.Info("Export started.");

        var manifestIndex = _manifestService.LoadIndex(manifestPath);

        using var session = new OutlookSession();
        session.Start();

        var pendingEntryIds = resumeState?.PendingEntryIds ?? BuildPendingList(session, options, log);
        var currentIndex = resumeState?.CurrentIndex ?? 0;

        summary.Total = pendingEntryIds.Count;

        for (int i = currentIndex; i < pendingEntryIds.Count; i++)
        {
            if (token.IsCancellationRequested)
            {
                summary.Canceled = true;
                summary.Processed = i;
                SaveState(statePath, options, exportRoot, pendingEntryIds, i);
                progress(new ProgressInfo { Processed = i, Total = pendingEntryIds.Count, Status = "Canceled. State saved." });
                log.Warn("Export canceled by user.");
                return summary;
            }

            var entryId = pendingEntryIds[i];
            progress(new ProgressInfo { Processed = i, Total = pendingEntryIds.Count, Status = $"Processing {i + 1} of {pendingEntryIds.Count}..." });

            try
            {
                var item = session.GetItemFromId(entryId, options.OutlookFolderStoreId);
                if (item is not MailItem mailItem)
                {
                    log.Warn($"Skipped non-mail item: {entryId}");
                    summary.Skipped++;
                    ComRelease.Release(item);
                    continue;
                }

                var messageClass = mailItem.MessageClass ?? string.Empty;
                if (!messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
                {
                    log.Warn($"Skipped MessageClass {messageClass}");
                    summary.Skipped++;
                    ComRelease.Release(mailItem);
                    continue;
                }

                var result = ExportMailItem(mailItem, options, emailsRoot, manifestPath, manifestIndex, log, summary);
                if (result == ExportStatus.Skipped)
                {
                    summary.Skipped++;
                }
                else if (result == ExportStatus.Failed)
                {
                    summary.Failed++;
                }
                else if (result == ExportStatus.Partial)
                {
                    summary.Partial++;
                }
                else
                {
                    summary.Processed++;
                }

                ComRelease.Release(mailItem);
            }
            catch (System.Exception ex)
            {
                summary.Failed++;
                log.Error($"Failed processing entry {entryId}: {ex.GetType().Name} - {ex.Message}. Stack trace: {ex.StackTrace}");
            }
        }

        _stateService.Delete(statePath);
        progress(new ProgressInfo { Processed = pendingEntryIds.Count, Total = pendingEntryIds.Count, Status = "Export complete." });
        log.Info("Export complete.");
        return summary;
    }

    private List<string> BuildPendingList(OutlookSession session, ExportOptions options, LogService log)
    {
        var pending = new List<string>();
        MAPIFolder? folder = null;
        Items? items = null;

        try
        {
            folder = session.GetFolder(options.OutlookFolderEntryId, options.OutlookFolderStoreId);
            items = folder.Items;
            items.SetColumns("EntryID,MessageClass,Subject,SenderEmailAddress,SentOn,ReceivedTime");

            object? item = items.GetFirst();
            while (item != null)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        pending.Add(mailItem.EntryID);
                    }
                    else
                    {
                        pending.Add(((dynamic)item).EntryID);
                    }
                }
                catch (Exception ex)
                {
                    log.Warn($"Failed to read item entry id: {ex.GetType().Name} - {ex.Message}");
                }
                finally
                {
                    ComRelease.Release(item);
                }

                item = items.GetNext();
            }
        }
        finally
        {
            ComRelease.Release(items);
            ComRelease.Release(folder);
        }

        return pending;
    }

    private ExportStatus ExportMailItem(MailItem mailItem, ExportOptions options, string emailsRoot, string manifestPath, Dictionary<string, ManifestEntry> manifestIndex, LogService log, ExportSummary summary)
    {
        var messageId = GetStringProperty(mailItem, "http://schemas.microsoft.com/mapi/proptag/0x1035001F");
        var searchKey = GetBinaryProperty(mailItem, "http://schemas.microsoft.com/mapi/proptag/0x300B0102");
        var sender = mailItem.SenderEmailAddress;
        var subject = mailItem.Subject;
        var sentUtc = ToUtc(mailItem.SentOn);
        var receivedUtc = ToUtc(mailItem.ReceivedTime);
        var timeForId = sentUtc ?? receivedUtc;
        var bodyPreview = TryGetBodyPreview(mailItem, log, out var bodyError);

        var entryIdBase = _idService.CreateDeterministicId(messageId, sender, subject, timeForId, bodyPreview, searchKey, mailItem.EntryID);
        var entryId = entryIdBase;

        var targetFolder = Path.Combine(emailsRoot, entryId);
        var msgPath = Path.Combine(targetFolder, $"{entryId}.msg");

        if (File.Exists(msgPath))
        {
            if (manifestIndex.TryGetValue(entryId, out var existing))
            {
                var existingHash = _hashService.ComputeSha256(msgPath);
                if (string.Equals(existing.Canonical.HashSha256, existingHash, StringComparison.OrdinalIgnoreCase))
                {
                    log.Info($"Skipped existing export for {entryId}");
                    return ExportStatus.Skipped;
                }
            }

            entryId = ResolveCollision(entryIdBase, emailsRoot, log, summary);
            targetFolder = Path.Combine(emailsRoot, entryId);
            msgPath = Path.Combine(targetFolder, $"{entryId}.msg");
        }

        Directory.CreateDirectory(targetFolder);
        var attachmentsFolder = Path.Combine(targetFolder, "attachments");
        Directory.CreateDirectory(attachmentsFolder);

        var headersPath = Path.Combine(targetFolder, $"{entryId}.headers.txt");
        var metaPath = Path.Combine(targetFolder, $"{entryId}_meta.json");

        var manifestEntry = new ManifestEntry
        {
            EntryId = entryId,
            ExportStatus = "complete",
            HeadersPath = PathHelper.ToRelativePath(Path.GetDirectoryName(manifestPath)!, headersPath),
            SourceMetadata = new EmailSourceMetadata
            {
                FolderPath = options.OutlookFolderPath,
                Subject = subject,
                Sender = sender,
                Recipients = ReadRecipients(mailItem),
                SentUtc = sentUtc?.ToString("o"),
                ReceivedUtc = receivedUtc?.ToString("o"),
                MessageId = messageId
            }
        };

        var attachments = new List<AttachmentRecord>();
        var status = ExportStatus.Complete;

        try
        {
            mailItem.SaveAs(msgPath, OlSaveAsType.olMSGUnicode);
            var msgHash = _hashService.ComputeSha256(msgPath);
            var msgSize = new FileInfo(msgPath).Length;

            manifestEntry.Canonical = new CanonicalRecord
            {
                Path = PathHelper.ToRelativePath(Path.GetDirectoryName(manifestPath)!, msgPath),
                HashSha256 = msgHash,
                SizeBytes = msgSize
            };

            var headers = GetStringProperty(mailItem, "http://schemas.microsoft.com/mapi/proptag/0x007D001F");
            File.WriteAllText(headersPath, headers ?? string.Empty);

                attachments.AddRange(ExtractAttachments(mailItem, attachmentsFolder, entryId, log, Path.GetDirectoryName(manifestPath)!));

            WriteMetadata(metaPath, manifestEntry, messageClass: mailItem.MessageClass, bodyError: bodyError);
        }
        catch (Exception ex)
        {
            status = ExportStatus.Failed;
            manifestEntry.ExportStatus = "failed";
            manifestEntry.Error = ex.Message;
            log.Error($"Export failed for {entryId}: {ex.GetType().Name} - {ex.Message}. Stack trace: {ex.StackTrace}");
        }

        if (bodyError != null && status != ExportStatus.Failed)
        {
            status = ExportStatus.Partial;
            manifestEntry.ExportStatus = "partial";
            manifestEntry.Error = bodyError;
            log.Warn($"Partial export for {entryId}: {bodyError}");
        }

        manifestEntry.Attachments = attachments;
        _manifestService.Append(manifestPath, manifestEntry);
        manifestIndex[manifestEntry.EntryId] = manifestEntry;
        return status;
    }

    private IEnumerable<AttachmentRecord> ExtractAttachments(MailItem mailItem, string attachmentsFolder, string entryId, LogService log, string exportRoot)
    {
        var results = new List<AttachmentRecord>();
        Attachments? attachments = null;

        try
        {
            attachments = mailItem.Attachments;
            var count = attachments.Count;
            for (int i = 1; i <= count; i++)
            {
                Attachment? attachment = null;
                try
                {
                    attachment = attachments[i];
                    var originalName = attachment.FileName;
                    var safeName = FileNameSanitizer.Sanitize(originalName, "attachment");
                    var prefixedName = $"{entryId}_att_{i:000}_{safeName}";
                    prefixedName = FileNameSanitizer.EnsureUnique(attachmentsFolder, prefixedName);
                    var path = Path.Combine(attachmentsFolder, prefixedName);

                    attachment.SaveAsFile(path);
                    var hash = _hashService.ComputeSha256(path);
                    var size = new FileInfo(path).Length;

                    results.Add(new AttachmentRecord
                    {
                        Path = PathHelper.ToRelativePath(exportRoot, path),
                        OriginalFilename = originalName ?? string.Empty,
                        HashSha256 = hash,
                        SizeBytes = size
                    });
                }
                catch (Exception ex)
                {
                    log.Warn($"Failed to extract attachment {i} for entry {entryId}: {ex.GetType().Name} - {ex.Message}");
                }
                finally
                {
                    ComRelease.Release(attachment);
                }
            }
        }
        finally
        {
            ComRelease.Release(attachments);
        }

        return results;
    }

    private static string? GetStringProperty(MailItem item, string daslTag)
    {
        PropertyAccessor? accessor = null;
        try
        {
            accessor = item.PropertyAccessor;
            return accessor.GetProperty(daslTag) as string;
        }
        catch
        {
            return null;
        }
        finally
        {
            ComRelease.Release(accessor);
        }
    }

    private static byte[]? GetBinaryProperty(MailItem item, string daslTag)
    {
        PropertyAccessor? accessor = null;
        try
        {
            accessor = item.PropertyAccessor;
            return accessor.GetProperty(daslTag) as byte[];
        }
        catch
        {
            return null;
        }
        finally
        {
            ComRelease.Release(accessor);
        }
    }

    private static DateTime? ToUtc(DateTime time)
    {
        if (time == DateTime.MinValue)
        {
            return null;
        }

        return time.Kind == DateTimeKind.Utc ? time : time.ToUniversalTime();
    }

    private static List<string> ReadRecipients(MailItem mailItem)
    {
        var recipients = new List<string>();
        Recipients? collection = null;

        try
        {
            collection = mailItem.Recipients;
            for (int i = 1; i <= collection.Count; i++)
            {
                Recipient? recipient = null;
                try
                {
                    recipient = collection[i];
                    var address = recipient.Address ?? recipient.Name;
                    if (!string.IsNullOrWhiteSpace(address))
                    {
                        recipients.Add(address);
                    }
                }
                finally
                {
                    ComRelease.Release(recipient);
                }
            }
        }
        finally
        {
            ComRelease.Release(collection);
        }

        return recipients;
    }

    private static string? TryGetBodyPreview(MailItem mailItem, LogService log, out string? error)
    {
        error = null;
        try
        {
            var body = mailItem.Body ?? string.Empty;
            return body.Length <= 512 ? body : body[..512];
        }
        catch (Exception ex)
        {
            error = $"Body inaccessible: {ex.Message}";
            log.Warn(error);
            return null;
        }
    }

    private void WriteMetadata(string metaPath, ManifestEntry entry, string? messageClass, string? bodyError)
    {
        var metadata = new
        {
            entry_id = entry.EntryId,
            message_class = messageClass,
            body_error = bodyError,
            source = entry.SourceMetadata
        };

        var json = System.Text.Json.JsonSerializer.Serialize(metadata, new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        });

        File.WriteAllText(metaPath, json);
    }

    private static string ResolveCollision(string baseEntryId, string emailsRoot, LogService log, ExportSummary summary)
    {
        const int maxRetries = 1000;
        var counter = 1;
        while (counter <= maxRetries)
        {
            var candidate = $"{baseEntryId}-collision-{counter}";
            var candidateFolder = Path.Combine(emailsRoot, candidate);
            if (!Directory.Exists(candidateFolder))
            {
                summary.Collisions++;
                log.Warn($"Collision detected for {baseEntryId}, using {candidate}");
                return candidate;
            }

            counter++;
        }

        throw new InvalidOperationException($"Unable to resolve collision for {baseEntryId} after {maxRetries} attempts");
    }

    private string ResolveExportRoot(ExportOptions options, out ExportState? resumeState)
    {
        resumeState = null;
        if (options.ResumeEnabled)
        {
            var statePath = FindLatestStatePath(options.OutputRoot);
            if (statePath != null)
            {
                resumeState = _stateService.Load(statePath);
                if (resumeState != null)
                {
                    return resumeState.ExportRoot;
                }
            }
        }

        var folder = $"Export_{DateTime.Now:yyyyMMdd_HHmmss}";
        return Path.Combine(options.OutputRoot, folder);
    }

    private void SaveState(string statePath, ExportOptions options, string exportRoot, List<string> pendingEntryIds, int currentIndex)
    {
        var state = new ExportState
        {
            RunId = Guid.NewGuid().ToString("N"),
            ExportRoot = exportRoot,
            FolderEntryId = options.OutlookFolderEntryId,
            FolderStoreId = options.OutlookFolderStoreId,
            FolderPath = options.OutlookFolderPath,
            CurrentIndex = currentIndex,
            PendingEntryIds = pendingEntryIds
        };

        _stateService.Save(statePath, state);
    }

    private static string? FindLatestStatePath(string outputRoot)
    {
        if (!Directory.Exists(outputRoot))
        {
            return null;
        }

        var stateFiles = Directory.GetFiles(outputRoot, "state.json", SearchOption.AllDirectories);
        if (stateFiles.Length == 0)
        {
            return null;
        }

        return stateFiles
            .Select(path => new FileInfo(path))
            .OrderByDescending(info => info.LastWriteTimeUtc)
            .FirstOrDefault()?.FullName;
    }

    private enum ExportStatus
    {
        Complete,
        Partial,
        Failed,
        Skipped
    }
}
