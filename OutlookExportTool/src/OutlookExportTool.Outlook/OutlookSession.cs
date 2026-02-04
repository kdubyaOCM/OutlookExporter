using Microsoft.Office.Interop.Outlook;

namespace OutlookExportTool.Outlook;

public sealed class OutlookSession : IDisposable
{
    private Application? _application;
    private NameSpace? _namespace;

    public Application Application => _application ?? throw new InvalidOperationException("Outlook session not started.");
    public NameSpace NameSpace => _namespace ?? throw new InvalidOperationException("Outlook session not started.");

    public void Start()
    {
        _application = new Application();
        _namespace = _application.GetNamespace("MAPI");
    }

    public MAPIFolder GetFolder(string entryId, string storeId)
    {
        return NameSpace.GetFolderFromID(entryId, storeId);
    }

    public object GetItemFromId(string entryId, string storeId)
    {
        return NameSpace.GetItemFromID(entryId, storeId);
    }

    public void Dispose()
    {
        if (_namespace != null)
        {
            ComRelease.Release(_namespace);
            _namespace = null;
        }

        if (_application != null)
        {
            ComRelease.Release(_application);
            _application = null;
        }
    }
}
