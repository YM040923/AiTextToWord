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

public sealed record ExportedFileActionState(string FilePath, string FileName, string FolderPath)
{
    public static ExportedFileActionState? FromFilePath(string? filePath)
    {
        var folderPath = ExportLocationMemory.ParentFolderFromFilePath(filePath);
        if (folderPath is null || string.IsNullOrWhiteSpace(filePath))
        {
            return null;
        }

        try
        {
            var fileName = Path.GetFileName(filePath);
            return string.IsNullOrWhiteSpace(fileName)
                ? null
                : new ExportedFileActionState(filePath, fileName, folderPath);
        }
        catch (ArgumentException)
        {
            return null;
        }
    }
}
