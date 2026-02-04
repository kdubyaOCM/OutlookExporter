# Outlook Export Tool (MVP-0)

Exports Classic Outlook (Microsoft 365) emails into deterministic evidence packs with MSG, headers, attachments, hashes, and JSONL manifest.

## Prerequisites
- Windows with Classic Outlook (Microsoft 365) installed and configured.
- .NET 8 SDK.
- Outlook must be able to access the target mailbox without prompts.

## Build
From `OutlookExportTool/`:

```bash
dotnet build OutlookExportTool.sln
```

## Run
```bash
dotnet run --project src/OutlookExportTool.WPF/OutlookExportTool.WPF.csproj
```

## Smoke Test (20 emails)
1. Launch the app.
2. Select a small Outlook folder containing about 20 emails.
3. Select an empty output folder.
4. Click Start and let it complete.

Expected output structure:
```
Export_YYYYMMDD_HHMMSS/
  manifest.jsonl
  logs/export_YYYYMMDD.log
  emails/{entry_id}/
    {entry_id}.msg
    {entry_id}.headers.txt
    {entry_id}_meta.json
    attachments/
      {entry_id}_att_001_{filename}
```

## Resume
If the export is canceled, `state.json` is written inside the export folder. Re-running with the same output folder resumes from `state.json` without re-scanning Outlook.

## Notes
- Only `IPM.Note*` message classes are exported.
- No PDF, OCR, or EML generation in MVP-0.
