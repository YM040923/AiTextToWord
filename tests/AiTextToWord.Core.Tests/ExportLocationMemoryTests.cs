using AiTextToWord.Core.Export;

namespace AiTextToWord.Core.Tests;

public sealed class ExportLocationMemoryTests
{
    [Fact]
    public void ParentFolderFromFilePath_ReturnsContainingFolder()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ai-text-to-word-export");
        var filePath = Path.Combine(folder, "AI文本导出.docx");

        var result = ExportLocationMemory.ParentFolderFromFilePath(filePath);

        Assert.Equal(folder, result);
    }

    [Fact]
    public void ParentFolderFromFilePath_IgnoresEmptyOrInvalidPaths()
    {
        Assert.Null(ExportLocationMemory.ParentFolderFromFilePath(null));
        Assert.Null(ExportLocationMemory.ParentFolderFromFilePath(string.Empty));
        Assert.Null(ExportLocationMemory.ParentFolderFromFilePath("AI文本导出.docx"));
    }

    [Fact]
    public void PickerSettingsIdentifier_IsStableForWindowsPickerMemory()
    {
        Assert.Equal("AiTextToWord.Export", ExportLocationMemory.PickerSettingsIdentifier);
    }
}
