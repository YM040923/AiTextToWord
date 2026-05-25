namespace AiTextToWord.Core.Export;

public static class ExportLocationMemory
{
    public const string PickerSettingsIdentifier = "AiTextToWord.Export";

    public static string? ParentFolderFromFilePath(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return null;
        }

        try
        {
            var folderPath = Path.GetDirectoryName(filePath);
            return string.IsNullOrWhiteSpace(folderPath) ? null : folderPath;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }
}
