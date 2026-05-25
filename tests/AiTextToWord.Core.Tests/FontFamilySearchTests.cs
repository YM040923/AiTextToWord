using AiTextToWord.Core.Fonts;

namespace AiTextToWord.Core.Tests;

public sealed class FontFamilySearchTests
{
    [Fact]
    public void Filter_ReturnsDistinctFontsWhenQueryIsEmpty()
    {
        var fonts = new[] { "微软雅黑", "Arial", "arial", "宋体" };

        var results = FontFamilySearch.Filter(fonts, string.Empty);

        Assert.Equal(["微软雅黑", "Arial", "宋体"], results);
    }

    [Fact]
    public void Filter_MatchesFontsByPartialName()
    {
        var fonts = new[] { "Microsoft YaHei", "微软雅黑", "SimSun", "Arial" };

        var results = FontFamilySearch.Filter(fonts, "ya");

        Assert.Equal(["Microsoft YaHei"], results);
    }

    [Fact]
    public void Filter_PrioritizesPrefixMatches()
    {
        var fonts = new[] { "Noto Sans", "Sans Serif", "Source Sans 3" };

        var results = FontFamilySearch.Filter(fonts, "sans");

        Assert.Equal(["Sans Serif", "Noto Sans", "Source Sans 3"], results);
    }
}
