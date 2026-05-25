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

    [Fact]
    public void ExportedFileActionState_FromFilePath_ProvidesFileNameAndFolder()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ai-text-to-word-export");
        var filePath = Path.Combine(folder, "AI文本导出.pdf");

        var state = ExportedFileActionState.FromFilePath(filePath);

        Assert.NotNull(state);
        Assert.Equal(filePath, state.FilePath);
        Assert.Equal("AI文本导出.pdf", state.FileName);
        Assert.Equal(folder, state.FolderPath);
    }

    [Fact]
    public void ExportedFileActionState_FromFilePath_IgnoresPathsWithoutFileNameOrFolder()
    {
        Assert.Null(ExportedFileActionState.FromFilePath(null));
        Assert.Null(ExportedFileActionState.FromFilePath(string.Empty));
        Assert.Null(ExportedFileActionState.FromFilePath("AI文本导出.docx"));
        Assert.Null(ExportedFileActionState.FromFilePath(Path.Combine(Path.GetTempPath(), "folder") + Path.DirectorySeparatorChar));
    }
}
