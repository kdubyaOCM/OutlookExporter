namespace OutlookExportTool.Core.Utilities;

public static class PathHelper
{
    public static string ToRelativePath(string root, string fullPath)
    {
        try
        {
            return Path.GetRelativePath(root, fullPath).Replace('\\', '/');
        }
        catch
        {
            return fullPath.Replace('\\', '/');
        }
    }
}
