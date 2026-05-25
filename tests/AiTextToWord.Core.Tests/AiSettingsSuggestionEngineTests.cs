using AiTextToWord.Core.Export;

namespace AiTextToWord.Core.Tests;

public sealed class AiSettingsSuggestionEngineTests
{
    [Fact]
    public void Suggest_CreatesFormalReportSettingsFromPlainLanguage()
    {
        var suggestion = AiSettingsSuggestionEngine.Suggest("我想要正式一点，适合发给老师，排版宽松一些", "微软雅黑");

        Assert.Equal(ExportFormat.Word, suggestion.Format);
        Assert.Equal(ExportPreset.FormalReport, suggestion.Preset);
        Assert.Equal(12, suggestion.BodyFontSize);
        Assert.Equal(22, suggestion.HeadingFontSize);
        Assert.Equal(1.5, suggestion.LineSpacing);
        Assert.Equal(ExportPageMargin.Wide, suggestion.PageMargin);
        Assert.Equal(ExportQuoteStyle.GrayBlock, suggestion.QuoteStyle);
        Assert.Equal(ExportListDensity.Comfortable, suggestion.ListDensity);
    }

    [Fact]
    public void Suggest_CreatesCompactSettingsFromPlainLanguage()
    {
        var suggestion = AiSettingsSuggestionEngine.Suggest("我只想自己看，尽量紧凑省纸", "微软雅黑");

        Assert.Equal(ExportPreset.CompactNotes, suggestion.Preset);
        Assert.Equal(10.5, suggestion.BodyFontSize);
        Assert.Equal(16, suggestion.HeadingFontSize);
        Assert.Equal(1.15, suggestion.LineSpacing);
        Assert.Equal(ExportPageMargin.Narrow, suggestion.PageMargin);
        Assert.Equal(ExportListDensity.Compact, suggestion.ListDensity);
    }

    [Fact]
    public void TryParseModelJson_UsesOnlyKnownSettingsAndClampsNumbers()
    {
        const string json = """
            {
              "format": "pdf",
              "preset": "formal",
              "fontFamily": "宋体",
              "bodyFontSize": 120,
              "headingFontSize": 6,
              "lineSpacing": 8,
              "pageMargin": "wide",
              "quoteStyle": "gray",
              "listDensity": "comfortable",
              "ignored": "value"
            }
            """;

        Assert.True(AiSettingsSuggestionEngine.TryParseModelJson(json, "微软雅黑", out var suggestion));
        Assert.Equal(ExportFormat.Pdf, suggestion.Format);
        Assert.Equal(ExportPreset.FormalReport, suggestion.Preset);
        Assert.Equal("宋体", suggestion.FontFamily);
        Assert.Equal(72, suggestion.BodyFontSize);
        Assert.Equal(10, suggestion.HeadingFontSize);
        Assert.Equal(3, suggestion.LineSpacing);
        Assert.Equal(ExportPageMargin.Wide, suggestion.PageMargin);
        Assert.Equal(ExportQuoteStyle.GrayBlock, suggestion.QuoteStyle);
        Assert.Equal(ExportListDensity.Comfortable, suggestion.ListDensity);
    }
}
