using System.Runtime.InteropServices;

namespace OutlookExportTool.Outlook;

public static class ComRelease
{
    public static void Release(object? comObject)
    {
        if (comObject == null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch
        {
            // Ignore release failures.
        }
    }
}
