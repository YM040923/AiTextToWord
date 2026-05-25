using AiTextToWord.Core.Fonts;

namespace AiTextToWord.Core.Tests;

public sealed class WindowsFontNameParserTests
{
    [Fact]
    public void ExtractFamilyNames_SplitsCombinedWindowsFontNames()
    {
        var names = WindowsFontNameParser.ExtractFamilyNames("微软雅黑 & Microsoft YaHei UI (TrueType)");

        Assert.Equal(["微软雅黑", "Microsoft YaHei UI"], names);
    }

    [Theory]
    [InlineData("Arial Bold Italic (TrueType)", "Arial")]
    [InlineData("Cascadia Code SemiLight (TrueType)", "Cascadia Code")]
    [InlineData("Microsoft YaHei UI Light (TrueType)", "Microsoft YaHei UI")]
    public void ExtractFamilyNames_RemovesCommonStyleSuffixes(string registryName, string expected)
    {
        var names = WindowsFontNameParser.ExtractFamilyNames(registryName);

        Assert.Equal([expected], names);
    }

    [Fact]
    public void ExtractFamilyNames_KeepsFamilyNamesThatContainStyleWords()
    {
        var names = WindowsFontNameParser.ExtractFamilyNames("Arial Black (TrueType)");

        Assert.Equal(["Arial Black"], names);
    }

    [Fact]
    public void ExtractFamilyNames_DropsVerticalFontAliases()
    {
        var names = WindowsFontNameParser.ExtractFamilyNames("@SimSun (TrueType)");

        Assert.Empty(names);
    }
}
